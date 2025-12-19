"""
说明：
【小红书评论数据抓取脚本】
在表格中填入要抓取的笔记链接，脚本会打开每个小红书链接，手动快速滚动所有评论，自动抓取所有评论数据并保存到CSV和JSON文件中。
用最简单的代码，实现最强不封禁的功能。

功能：
1. 启动Chrome浏览器，进入小红书首页，等待用户扫码登录（登录信息会保存）
2. 从Excel文件或变量列表读取多个视频链接
3. 在新标签页打开每个链接
4. 监听评论API请求，自动解析响应体数据
5. 保存数据到CSV和JSON文件（按note_id分别保存，同时保存总文件）

使用方法：
1. 安装依赖：pip install playwright openpyxl rich
2. 安装浏览器：playwright install chromium
3. 在URL_LIST变量中填入链接，或准备xiaohongshuVidioUrlTemplate.xlsx文件
4. 运行脚本：python xhs_manusCrawlerComments.py
5. 在浏览器中扫码登录（首次运行）
6. 脚本会自动打开所有链接并监听评论数据
7. 当用户在浏览器中翻页查看评论时，数据会自动保存
8. 按Ctrl+C停止脚本

输出文件：
- 每个note_id会生成：1_[标题]_[note_id].json 和 1_[标题]_[note_id].csv
- 所有数据汇总：All CommentData.json 和 All CommentData.csv

注意：
- 首次运行需要登录，登录信息会保存，下次运行无需重新登录
- 脚本会持续运行，监听评论数据，直到用户手动停止
- 需要确保Excel文件格式正确（第一行为标题，从第二行开始是URL）

待办：
1、做各种平台评论数据抓取的适配模板，代码中配置平台的评论接口url、响应体、post id 这3项数据即可适配相应的平台。就你跑代码抓取不同平台的评论数据了，而且这种方式永远不会被平台认为是自动化爬数据。
"""

import asyncio
import csv
import json
import re
import os
from pathlib import Path
from playwright.async_api import async_playwright, BrowserContext, Page, Response
from typing import Dict, List, Set
import openpyxl
from rich.console import Console
from rich.live import Live
from rich.table import Table
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
from rich.layout import Layout
from rich.text import Text
from rich import box


# ========== 配置区域：可以在这里填入要爬取的小红书笔记链接 ==========
URL_LIST = [
    # 'http://xhslink.com/o/315a7XzU2Ho',
    # 'https://www.xiaohongshu.com/user/profile/59f8c90c11be103e7447cb0d?xsec_token=AB077S8P4Fy8-0QLGd0-QTJ0trZXaqrP9Q6WmBrS1Gvts=&xsec_source=pc_note',
]
# ================================================================

# Excel文件路径（相对于脚本文件所在目录）
SCRIPT_DIR = Path(__file__).parent.absolute()
EXCEL_FILE = str(SCRIPT_DIR / 'xiaohongshuVidioUrlTemplate.xlsx')

# 数据文件保存目录（与脚本同级）
DATA_FILE_DIR = SCRIPT_DIR / 'DataFile'

# 评论API URL模式
COMMENT_API_PATTERN = 'https://edith.xiaohongshu.com/api/sns/web/v2/comment/page?note_id='

# 存储所有评论数据的字典，key为note_id，value为嵌套结构的评论列表（一级评论包含sub_comments）
all_comments_data: Dict[str, List[Dict]] = {}

# 存储扁平化的评论数据（用于CSV导出），key为note_id
all_comments_flat: Dict[str, List[Dict]] = {}

# 存储已处理的note_id，避免重复处理
processed_note_ids: Set[str] = set()

# 存储note_id和页面标题的映射关系
note_id_to_title: Dict[str, str] = {}

# 存储note_id和页面的映射关系
note_id_to_page: Dict[str, Page] = {}

# 存储note_id和文件序号的映射关系
note_id_to_index: Dict[str, int] = {}

# 存储note_id和总评论数的映射关系
note_id_to_total_count: Dict[str, int] = {}

# Rich 控制台对象
console = Console()

# 全局 Live 对象（用于实时更新显示）
live_display = None


def get_user_data_dir():
    """获取持久化用户数据目录（保存登录状态）"""
    home_dir = Path.home()
    data_dir = home_dir / '.playwright-xhs-crawler'
    data_dir.mkdir(exist_ok=True)
    return str(data_dir)


def ensure_data_file_dir():
    """确保数据文件目录存在，如果不存在则创建"""
    DATA_FILE_DIR.mkdir(exist_ok=True)
    return str(DATA_FILE_DIR)


def read_urls_from_excel(file_path: str) -> List[str]:
    """从Excel文件读取URL列表（忽略首行标题）"""
    urls = []
    try:
        if os.path.exists(file_path):
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            
            # 从第二行开始读取（忽略首行标题）
            for row in ws.iter_rows(min_row=2, values_only=True):
                for cell_value in row:
                    if cell_value and isinstance(cell_value, str):
                        url = str(cell_value).strip()
                        if url and (url.startswith('http://') or url.startswith('https://')):
                            urls.append(url)
            
            wb.close()
            console.print(f'[green]✓[/green] 从Excel文件读取到 [cyan]{len(urls)}[/cyan] 个URL')
        else:
            console.print(f'[yellow]⚠[/yellow] Excel文件不存在: {file_path}')
    except Exception as e:
        console.print(f'[yellow]⚠[/yellow] 读取Excel文件失败: {str(e)}')
    
    return urls


def extract_note_id_from_url(url: str) -> str:
    """从URL中提取note_id"""
    # 匹配 /explore/ 后面的note_id
    match = re.search(r'/explore/([a-f0-9]+)', url)
    if match:
        return match.group(1)
    
    # 匹配 note_id= 参数
    match = re.search(r'note_id=([a-f0-9]+)', url)
    if match:
        return match.group(1)
    
    return ''


def parse_comment_response(response_data: dict, note_id: str) -> tuple:
    """
    解析评论API响应体数据
    返回：(嵌套结构的评论列表, 扁平化的评论列表)
    """
    nested_comments = []
    flat_comments = []
    
    try:
        if not response_data.get('success') or response_data.get('code') != 0:
            return nested_comments, flat_comments
        
        data = response_data.get('data', {})
        comments_list = data.get('comments', [])
        
        for comment in comments_list:
            # 解析一级评论（嵌套结构，用于JSON）
            comment_data = {
                'content': comment.get('content', ''),
                'like_count': comment.get('like_count', '0'),
                'ip_location': comment.get('ip_location', ''),
                'nickname': comment.get('user_info', {}).get('nickname', ''),
                'comment_id': comment.get('id', ''),
                'sub_comments': []  # 二级评论嵌套在这里
            }
            
            # 扁平化的一级评论（用于CSV）
            flat_comment = {
                'content': comment.get('content', ''),
                'like_count': comment.get('like_count', '0'),
                'ip_location': comment.get('ip_location', ''),
                'nickname': comment.get('user_info', {}).get('nickname', ''),
                'note_id': note_id,
                'comment_id': comment.get('id', ''),
                'parent_comment_id': '',
                'is_sub_comment': False
            }
            flat_comments.append(flat_comment)
            
            # 解析二级评论（sub_comments）
            sub_comments = comment.get('sub_comments', [])
            for sub_comment in sub_comments:
                # 嵌套结构的二级评论（用于JSON）
                sub_comment_data = {
                    'content': sub_comment.get('content', ''),
                    'like_count': sub_comment.get('like_count', '0'),
                    'ip_location': sub_comment.get('ip_location', ''),
                    'nickname': sub_comment.get('user_info', {}).get('nickname', ''),
                    'comment_id': sub_comment.get('id', '')
                }
                comment_data['sub_comments'].append(sub_comment_data)
                
                # 扁平化的二级评论（用于CSV）
                flat_sub_comment = {
                    'content': sub_comment.get('content', ''),
                    'like_count': sub_comment.get('like_count', '0'),
                    'ip_location': sub_comment.get('ip_location', ''),
                    'nickname': sub_comment.get('user_info', {}).get('nickname', ''),
                    'note_id': note_id,
                    'comment_id': sub_comment.get('id', ''),
                    'parent_comment_id': comment.get('id', ''),
                    'is_sub_comment': True
                }
                flat_comments.append(flat_sub_comment)
            
            nested_comments.append(comment_data)
        
        total_count = len(flat_comments)
        # 使用 rich 更新显示
        update_display()
        
    except Exception as e:
        console.print(f'  [yellow]⚠[/yellow] 解析评论数据失败: {str(e)}')
    
    return nested_comments, flat_comments


async def handle_comment_api_response(response: Response):
    """处理评论API响应，解析响应体数据"""
    try:
        url = response.url
        
        # 检查是否是评论API
        if COMMENT_API_PATTERN not in url:
            return
        
        # 从URL中提取note_id
        note_id_match = re.search(r'note_id=([a-f0-9]+)', url)
        if not note_id_match:
            return
        
        note_id = note_id_match.group(1)
        
        # 获取响应体
        try:
            response_body = await response.json()
        except:
            # 如果响应不是JSON，忽略
            return
        
        # 解析评论数据
        nested_comments, flat_comments = parse_comment_response(response_body, note_id)
        
        if nested_comments or flat_comments:
            # 初始化数据结构
            if note_id not in all_comments_data:
                all_comments_data[note_id] = []
            if note_id not in all_comments_flat:
                all_comments_flat[note_id] = []
            
            # 收集已存在的comment_id（嵌套结构）
            existing_nested_ids = set()
            for c in all_comments_data[note_id]:
                existing_nested_ids.add(c['comment_id'])
                for sc in c.get('sub_comments', []):
                    existing_nested_ids.add(sc['comment_id'])
            
            # 收集已存在的comment_id（扁平结构）
            existing_flat_ids = {c['comment_id'] for c in all_comments_flat[note_id]}
            
            # 过滤新评论
            new_nested_comments = []
            new_flat_comments = []
            
            for nested_c in nested_comments:
                # 检查一级评论是否已存在
                if nested_c['comment_id'] not in existing_nested_ids:
                    new_nested_comments.append(nested_c)
                    existing_nested_ids.add(nested_c['comment_id'])
                    
                    # 找到对应的一级评论扁平数据
                    flat_c = next((fc for fc in flat_comments if fc['comment_id'] == nested_c['comment_id'] and not fc.get('is_sub_comment', False)), None)
                    if flat_c and flat_c['comment_id'] not in existing_flat_ids:
                        new_flat_comments.append(flat_c)
                        existing_flat_ids.add(flat_c['comment_id'])
                    
                    # 处理二级评论
                    for sub_c in nested_c.get('sub_comments', []):
                        if sub_c['comment_id'] not in existing_nested_ids:
                            existing_nested_ids.add(sub_c['comment_id'])
                            # 找到对应的二级评论扁平数据
                            flat_sub_c = next((fc for fc in flat_comments if fc['comment_id'] == sub_c['comment_id'] and fc.get('is_sub_comment', False)), None)
                            if flat_sub_c and flat_sub_c['comment_id'] not in existing_flat_ids:
                                new_flat_comments.append(flat_sub_c)
                                existing_flat_ids.add(flat_sub_c['comment_id'])
            
            if new_nested_comments:
                all_comments_data[note_id].extend(new_nested_comments)
                all_comments_flat[note_id].extend(new_flat_comments)
                # 使用 rich 更新显示
                update_display()
                
                # 尝试获取页面标题和总评论数（如果还没有）
                if note_id not in note_id_to_title or note_id not in note_id_to_total_count:
                    # 尝试从当前响应的页面获取
                    try:
                        page = response.request.frame.page if hasattr(response.request, 'frame') else None
                        if page and not page.is_closed():
                            if note_id not in note_id_to_title:
                                title = await get_page_title(page)
                                if title:
                                    note_id_to_title[note_id] = title
                            if note_id not in note_id_to_total_count:
                                total_count = await get_total_comment_count(page)
                                if total_count > 0:
                                    note_id_to_total_count[note_id] = total_count
                        else:
                            # 查找包含该note_id的页面
                            for pid, p in note_id_to_page.items():
                                if pid == note_id and not p.is_closed():
                                    if note_id not in note_id_to_title:
                                        title = await get_page_title(p)
                                        if title:
                                            note_id_to_title[note_id] = title
                                    if note_id not in note_id_to_total_count:
                                        total_count = await get_total_comment_count(p)
                                        if total_count > 0:
                                            note_id_to_total_count[note_id] = total_count
                                    break
                    except:
                        pass
                
                # 保存数据到文件
                await save_comments_to_file(note_id)
                await save_all_comments_to_file()
        
    except Exception as e:
        console.print(f'  [yellow]⚠[/yellow] 处理评论API响应失败: {str(e)}')


def create_status_panel() -> Panel:
    """创建实时状态显示面板"""
    # 计算总评论数
    total_comments = sum(len(comments) for comments in all_comments_flat.values())
    total_note_ids = len(all_comments_data)
    
    # 创建统计信息表格
    stats_table = Table(show_header=False, box=None, padding=(0, 1))
    stats_table.add_row("📊 总评论数:", f"[bold green]{total_comments}[/bold green]")
    stats_table.add_row("📝 处理笔记数:", f"[bold cyan]{total_note_ids}[/bold cyan]")
    
    # 创建每个note_id的详细表格
    if all_comments_data:
        detail_table = Table(title="📋 各笔记评论统计", box=box.ROUNDED, show_header=True, header_style="bold magenta")
        detail_table.add_column("序号", style="cyan", width=6)
        detail_table.add_column("Note ID", style="yellow", width=20)
        detail_table.add_column("标题", style="green", width=30, overflow="ellipsis")
        detail_table.add_column("评论数", style="bold blue", justify="right", width=12)
        detail_table.add_column("进度", style="bold", justify="right", width=15)
        
        for note_id, comments in sorted(all_comments_data.items(), 
                                       key=lambda x: note_id_to_index.get(x[0], 999)):
            index = note_id_to_index.get(note_id, 0)
            title = note_id_to_title.get(note_id, '未知标题')
            count = len(all_comments_flat.get(note_id, []))
            total_count = note_id_to_total_count.get(note_id, 0)
            
            # 计算进度
            if total_count > 0:
                progress_percent = (count / total_count) * 100
                if progress_percent >= 100:
                    progress_text = f"[bold green]✓ 100%[/bold green]"
                else:
                    progress_text = f"[yellow]{progress_percent:.1f}%[/yellow]"
            else:
                progress_text = "[dim]等待中...[/dim]"
            
            detail_table.add_row(
                str(index),
                note_id[:16] + "..." if len(note_id) > 16 else note_id,
                title[:28] + "..." if len(title) > 28 else title,
                f"{count}/{total_count}" if total_count > 0 else str(count),
                progress_text
            )
    else:
        detail_table = Table(title="📋 各笔记评论统计", box=box.ROUNDED)
        detail_table.add_column("状态", style="yellow")
        detail_table.add_row("等待数据...")
    
    # 创建布局
    layout = Layout()
    layout.split_column(
        Layout(Panel(stats_table, title="📈 总体统计", border_style="green"), size=5),
        Layout(detail_table)
    )
    
    return Panel(layout, title="[bold blue]小红书评论抓取实时监控[/bold blue]", border_style="blue")


def update_display():
    """更新实时显示"""
    global live_display
    if live_display:
        live_display.update(create_status_panel())


def sanitize_filename(filename: str) -> str:
    """清理文件名，移除非法字符"""
    # 移除或替换非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    filename = re.sub(illegal_chars, '_', filename)
    # 限制文件名长度
    if len(filename) > 200:
        filename = filename[:200]
    return filename


async def get_page_title(page: Page) -> str:
    """获取页面标题"""
    try:
        title = await page.title()
        return title.strip()
    except:
        return ''


async def get_total_comment_count(page: Page) -> int:
    """从页面获取评论总数（通过class="total"元素）"""
    try:
        # 查找 class="total" 的元素
        total_element = await page.query_selector('.total')
        if total_element:
            text = await total_element.inner_text()
            # 提取数字，例如 "共 92 条评论" -> 92
            match = re.search(r'共\s*(\d+)\s*条评论', text)
            if match:
                return int(match.group(1))
    except Exception as e:
        # 静默处理错误
        pass
    return 0


async def save_comments_to_file(note_id: str):
    """保存指定note_id的评论数据到文件"""
    if note_id not in all_comments_data or not all_comments_data[note_id]:
        return
    
    nested_comments = all_comments_data[note_id]
    flat_comments = all_comments_flat.get(note_id, [])
    
    # 获取页面标题
    page_title = note_id_to_title.get(note_id, '')
    
    # 如果还没有标题，尝试从页面获取
    if not page_title and note_id in note_id_to_page:
        try:
            page = note_id_to_page[note_id]
            if not page.is_closed():
                page_title = await get_page_title(page)
                note_id_to_title[note_id] = page_title
        except:
            pass
    
    # 如果还是没有标题，使用默认值
    if not page_title:
        page_title = f'note_{note_id}'
    
    # 获取文件序号
    file_index = note_id_to_index.get(note_id, 1)
    
    # 生成文件名：序号+标题:+title+note_id值:+note_id
    safe_title = sanitize_filename(page_title)
    filename_base = f'{file_index} 标题:{safe_title} note_id值:{note_id}'
    
    # 确保数据文件目录存在
    ensure_data_file_dir()
    
    # 保存JSON（嵌套结构）
    json_path = DATA_FILE_DIR / f'{filename_base}.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(nested_comments, f, ensure_ascii=False, indent=2)
    
    # 保存CSV（扁平结构，按指定列顺序）
    csv_path = DATA_FILE_DIR / f'{filename_base}.csv'
    if flat_comments:
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'content', 'like_count', 'ip_location', 'nickname', 
                'note_id', 'comment_id', 'parent_comment_id', 'is_sub_comment'
            ])
            writer.writeheader()
            writer.writerows(flat_comments)
    
    # 文件保存信息通过 rich 显示，这里不单独打印
    update_display()


async def save_all_comments_to_file():
    """保存所有评论数据到总文件"""
    if not all_comments_data:
        return
    
    # 合并所有嵌套评论（用于JSON）
    all_nested_comments = []
    for note_id, nested_comments in all_comments_data.items():
        all_nested_comments.extend(nested_comments)
    
    # 合并所有扁平评论（用于CSV）
    all_flat_comments = []
    for note_id, flat_comments in all_comments_flat.items():
        all_flat_comments.extend(flat_comments)
    
    if not all_nested_comments and not all_flat_comments:
        return
    
    # 确保数据文件目录存在
    ensure_data_file_dir()
    
    # 保存JSON（嵌套结构）
    json_path = DATA_FILE_DIR / 'All CommentData.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(all_nested_comments, f, ensure_ascii=False, indent=2)
    
    # 保存CSV（扁平结构，按指定列顺序）
    csv_path = DATA_FILE_DIR / 'All CommentData.csv'
    if all_flat_comments:
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'content', 'like_count', 'ip_location', 'nickname', 
                'note_id', 'comment_id', 'parent_comment_id', 'is_sub_comment'
            ])
            writer.writeheader()
            writer.writerows(all_flat_comments)
    
    # 计算所有note_id的总评论数
    total_all = sum(len(comments) for comments in all_comments_flat.values())
    # 使用 rich 更新显示
    update_display()


def setup_response_listener(context: BrowserContext):
    """设置响应监听器，监听评论API响应"""
    async def response_handler(response: Response):
        await handle_comment_api_response(response)
    
    context.on('response', response_handler)
    console.print('[green]✓[/green] 已设置评论API响应监听器')


async def open_url_in_new_tab(context: BrowserContext, url: str, index: int, total: int):
    """在新标签页打开URL"""
    try:
        page = await context.new_page()
        
        console.print(f'\n[cyan][{index + 1}/{total}][/cyan] 打开链接: [dim]{url}[/dim]')
        
        # 导航到URL
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        
        # 等待页面加载
        await page.wait_for_timeout(3000)
        
        # 获取页面标题和note_id
        title = await page.title()
        note_id = extract_note_id_from_url(url)
        
        if not note_id:
            note_id = extract_note_id_from_url(page.url)
        
        console.print(f'  [green]页面标题:[/green] {title}')
        console.print(f'  [green]note_id:[/green] {note_id}')
        
        # 获取页面总评论数
        total_count = await get_total_comment_count(page)
        if total_count > 0:
            console.print(f'  [green]总评论数:[/green] {total_count}')
        
        # 保存note_id和标题的映射关系
        if note_id:
            note_id_to_title[note_id] = title
            note_id_to_page[note_id] = page
            note_id_to_index[note_id] = index + 1  # 记录文件序号（从1开始）
            if total_count > 0:
                note_id_to_total_count[note_id] = total_count
            processed_note_ids.add(note_id)
            # 如果已经有评论数据，立即保存
            if note_id in all_comments_data:
                await save_comments_to_file(note_id)
            # 更新显示
            update_display()
        
        return page
        
    except Exception as e:
        console.print(f'  [red]✗[/red] 打开链接失败: {str(e)}')
        return None


async def wait_for_login(page: Page):
    """等待用户登录"""
    console.print('\n[yellow]等待用户扫码登录小红书...[/yellow]')
    console.print('[dim]请在浏览器中完成登录，登录完成后脚本会自动继续[/dim]')
    
    max_wait_time = 300  # 最多等待5分钟
    check_interval = 2  # 每2秒检查一次
    elapsed_time = 0
    
    while elapsed_time < max_wait_time:
        try:
            # 检查是否已登录（通过检查页面内容）
            is_logged_in = await page.evaluate("""
                () => {
                    const text = document.body.innerText || '';
                    return !text.includes('登录') && 
                           !text.includes('立即登录') && 
                           !text.includes('请登录') &&
                           !text.includes('扫码登录');
                }
            """)
            
            if is_logged_in:
                console.print('[green]✓[/green] 检测到已登录，继续执行...')
                await page.wait_for_timeout(2000)
                return True
            
            await asyncio.sleep(check_interval)
            elapsed_time += check_interval
            
            if elapsed_time % 10 == 0:
                console.print(f'  [dim]等待中... ({elapsed_time}/{max_wait_time}秒)[/dim]')
                
        except Exception as e:
            console.print(f'  [yellow]检查登录状态时出错: {str(e)}[/yellow]')
            await asyncio.sleep(check_interval)
            elapsed_time += check_interval
    
    console.print('[yellow]⚠[/yellow] 等待登录超时，继续执行...')
    return False


async def main():
    """主函数"""
    # 读取URL列表
    urls = []
    
    # 优先从变量列表读取
    if URL_LIST:
        urls.extend(URL_LIST)
        console.print(f'[green]✓[/green] 从变量列表读取到 [cyan]{len(URL_LIST)}[/cyan] 个URL')
    
    # 从Excel文件读取
    excel_urls = read_urls_from_excel(EXCEL_FILE)
    urls.extend(excel_urls)
    
    # 去重
    urls = list(dict.fromkeys(urls))  # 保持顺序的去重
    
    if not urls:
        console.print('[red]错误：未找到任何URL，请在URL_LIST变量中填入链接或确保Excel文件存在且包含URL[/red]')
        return
    
    console.print(f'\n[bold cyan]总共需要处理的URL数量: {len(urls)}[/bold cyan]')
    
    user_data_dir = get_user_data_dir()
    
    console.print('\n[bold blue]启动浏览器（使用持久化上下文，登录状态会被保存）...[/bold blue]')
    console.print(f'[dim]用户数据目录: {user_data_dir}[/dim]')
    
    async with async_playwright() as p:
        # 使用 launch_persistent_context 创建持久化上下文
        context = await p.chromium.launch_persistent_context(
            user_data_dir,
            headless=False,  # 显示浏览器窗口
            channel='chrome',  # 使用系统安装的 Chrome
            args=[
                '--disable-blink-features=AutomationControlled',  # 隐藏自动化特征
            ],
        )
        
        # 设置页面默认缩放为80%
        await context.add_init_script("""
            (function() {
                function setZoom() {
                    if (document.body) {
                        document.body.style.zoom = '0.8';
                    }
                    if (document.documentElement) {
                        document.documentElement.style.zoom = '0.8';
                    }
                }
                // 立即设置
                setZoom();
                // 监听 DOM 变化，确保缩放生效
                const observer = new MutationObserver(setZoom);
                if (document.body) {
                    observer.observe(document.body, { attributes: true, attributeFilter: ['style'] });
                }
                observer.observe(document.documentElement, { 
                    childList: true, 
                    subtree: true,
                    attributes: true,
                    attributeFilter: ['style']
                });
                // 页面加载完成后再次设置
                if (document.readyState === 'complete' || document.readyState === 'interactive') {
                    setZoom();
                } else {
                    window.addEventListener('load', setZoom);
                    document.addEventListener('DOMContentLoaded', setZoom);
                }
            })();
        """)
        
        try:
            # 设置响应监听器
            setup_response_listener(context)
            
            # 打开小红书首页
            console.print('\n[bold blue]打开小红书首页...[/bold blue]')
            home_page = await context.new_page()
            await home_page.goto('https://www.xiaohongshu.com/', wait_until='domcontentloaded', timeout=60000)
            await home_page.wait_for_timeout(3000)
            
            # 等待用户登录
            await wait_for_login(home_page)
            
            console.print('\n[bold blue]开始打开视频链接...[/bold blue]')
            pages = []
            
            # 在新标签页打开每个URL
            for i, url in enumerate(urls):
                page = await open_url_in_new_tab(context, url, i, len(urls))
                if page:
                    pages.append(page)
                    # 记录note_id和文件序号的映射（从1开始）
                    note_id = extract_note_id_from_url(url)
                    if not note_id:
                        note_id = extract_note_id_from_url(page.url)
                    if note_id and note_id not in note_id_to_index:
                        note_id_to_index[note_id] = i + 1
                
                # 如果不是最后一个，等待一下
                if i < len(urls) - 1:
                    await asyncio.sleep(2)
            
            console.print(f'\n[bold green]✓[/bold green] 已打开 [bold cyan]{len(pages)}[/bold cyan] 个标签页')
            console.print('\n[bold yellow]开始监听评论API请求...[/bold yellow]')
            console.print('[dim]脚本将持续运行，监听评论数据[/dim]')
            console.print('[dim]当用户在浏览器中翻页查看评论时，数据会自动保存[/dim]')
            console.print('[dim]按Ctrl+C停止脚本运行[/dim]\n')
            
            # 使用 Live 实时更新显示
            global live_display
            with Live(create_status_panel(), refresh_per_second=2, screen=False) as live:
                live_display = live
                
                # 持续运行，直到用户中断
                try:
                    check_count = 0
                    while True:
                        await asyncio.sleep(2)  # 更频繁地更新显示
                        # 定期保存数据
                        if all_comments_data:
                            await save_all_comments_to_file()
                        
                        # 每10次循环（约20秒）检查一次总评论数
                        check_count += 1
                        if check_count >= 10:
                            check_count = 0
                            # 尝试更新总评论数
                            for note_id, page in note_id_to_page.items():
                                if note_id not in note_id_to_total_count or note_id_to_total_count[note_id] == 0:
                                    try:
                                        if not page.is_closed():
                                            total_count = await get_total_comment_count(page)
                                            if total_count > 0:
                                                note_id_to_total_count[note_id] = total_count
                                    except:
                                        pass
                        
                        # 更新显示
                        update_display()
                except (KeyboardInterrupt, asyncio.CancelledError):
                    console.print('\n\n[yellow]用户中断程序，正在保存数据...[/yellow]')
                finally:
                    live_display = None
            
            # 保存最终数据
            if all_comments_data:
                console.print('\n[yellow]保存最终数据...[/yellow]')
                try:
                    for note_id in all_comments_data.keys():
                        await save_comments_to_file(note_id)
                    await save_all_comments_to_file()
                    
                    total_comments = sum(len(comments) for comments in all_comments_flat.values())
                    console.print(f'\n[bold green]✓[/bold green] 数据保存完成！')
                    console.print(f'  [cyan]共处理 {len(all_comments_data)} 个note_id[/cyan]')
                    console.print(f'  [cyan]共抓取 {total_comments} 条评论[/cyan]')
                    console.print(f'  [dim]数据已保存到对应的JSON和CSV文件中[/dim]')
                except Exception:
                    pass  # 静默处理保存数据时的异常
            else:
                console.print('\n[yellow]⚠[/yellow] 未抓取到任何评论数据')
            
        except (KeyboardInterrupt, asyncio.CancelledError):
            # 用户中断，不显示错误信息
            pass
        except Exception as err:
            # 其他异常才显示错误信息
            console.print(f'\n[red]运行出错：{err}[/red]')
            import traceback
            traceback.print_exc()
        finally:
            # 优雅关闭浏览器，捕获所有可能的异常
            try:
                await context.close()
            except Exception:
                # 静默处理关闭浏览器时的异常（如连接已关闭）
                pass
            console.print('\n\n[dim]已关闭浏览器...[/dim]\n')


if __name__ == '__main__':
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, asyncio.CancelledError):
        # 用户中断，不显示错误信息
        console.print('\n[dim]已关闭浏览器[/dim]')
        exit(0)
    except Exception as err:
        # 其他异常才显示错误信息
        console.print(f'[red]运行出错：{err}[/red]')
        import traceback
        traceback.print_exc()
        exit(1)

