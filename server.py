#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
MCP Docx处理服务
提供docx文档的各种操作，包括查询、添加、修改、删除，以及字体样式设置等
使用官方MCP库实现
"""

import os
import tempfile
import logging
import traceback
from contextlib import asynccontextmanager
from typing import AsyncIterator, Dict, Any, Optional

from mcp.server.fastmcp import FastMCP, Context
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(tempfile.gettempdir(), "docx_mcp_server.log")),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("DocxMCPServer")

# 创建一个状态文件，用于在MCP服务重启时恢复状态
CURRENT_DOC_FILE = os.path.join(tempfile.gettempdir(), "docx_mcp_current_doc.txt")

class DocxProcessor:
    """处理Docx文档的类，实现各种文档操作"""
    
    def __init__(self):
        self.documents = {}  # 存储已打开的文档
        self.current_document = None
        self.current_file_path = None
        
        # 尝试从标记文件加载当前文档
        self._load_current_document()
    
    def _load_current_document(self):
        """从标记文件加载当前文档"""
        if not os.path.exists(CURRENT_DOC_FILE):
            return False
        
        try:
            with open(CURRENT_DOC_FILE, 'r', encoding='utf-8') as f:
                file_path = f.read().strip()
            
            if file_path and os.path.exists(file_path):
                self.current_file_path = file_path
                self.current_document = Document(file_path)
                self.documents[file_path] = self.current_document
                return True
        except Exception as e:
            logger.error(f"加载当前文档失败: {e}")
        
        return False
    
    def _save_current_document(self):
        """保存当前文档路径到标记文件"""
        if not self.current_file_path:
            return False
        
        try:
            with open(CURRENT_DOC_FILE, 'w', encoding='utf-8') as f:
                f.write(self.current_file_path)
            return True
        except Exception as e:
            logger.error(f"保存当前文档路径失败: {e}")
        
        return False
    
    def save_state(self):
        """保存处理器状态"""
        # 保存当前文档
        if self.current_document and self.current_file_path:
            try:
                self.current_document.save(self.current_file_path)
                self._save_current_document()
            except Exception as e:
                logger.error(f"保存当前文档失败: {e}")
    
    def load_state(self):
        """加载处理器状态"""
        self._load_current_document()

    # ... 保留原有的所有文档处理方法 ...

# 创建全局的处理器实例
processor = DocxProcessor()

@asynccontextmanager
async def server_lifespan(server: FastMCP) -> AsyncIterator[Dict[str, Any]]:
    """管理服务器生命周期"""
    try:
        # 服务器启动时加载状态
        logger.info("DocxProcessor MCP服务器启动中...")
        processor.load_state()
        yield {"processor": processor}
    finally:
        # 服务器关闭时保存状态
        logger.info("DocxProcessor MCP服务器关闭中...")
        processor.save_state()

# 创建MCP服务器
mcp = FastMCP(
    name="DocxProcessor",
    instructions="Word文档处理服务，提供创建、编辑、查询文档的功能",
    lifespan=server_lifespan
)

@mcp.tool()
def create_document(ctx: Context, file_path: str) -> str:
    """
    创建一个新的Word文档
    
    Parameters:
    - file_path: 文档保存路径
    """
    try:
        processor.current_document = Document()
        processor.current_file_path = file_path
        processor.documents[file_path] = processor.current_document
        
        # 保存文档
        processor.current_document.save(file_path)
        
        return f"成功创建文档: {file_path}"
    except Exception as e:
        error_msg = f"创建文档失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def open_document(ctx: Context, file_path: str) -> str:
    """
    打开一个现有的Word文档
    
    Parameters:
    - file_path: 要打开的文档路径
    """
    try:
        if not os.path.exists(file_path):
            return f"文件不存在: {file_path}"
        
        processor.current_document = Document(file_path)
        processor.current_file_path = file_path
        processor.documents[file_path] = processor.current_document
        
        return f"成功打开文档: {file_path}"
    except Exception as e:
        error_msg = f"打开文档失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def save_document(ctx: Context) -> str:
    """
    保存当前打开的Word文档到原文件（更新原文件）
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        if not processor.current_file_path:
            return "当前文档未保存过，请使用save_as_document指定保存路径"
            
        # 保存到原文件路径
        processor.current_document.save(processor.current_file_path)
        
        return f"成功保存文档到原文件: {processor.current_file_path}"
    except Exception as e:
        error_msg = f"保存文档失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def add_paragraph(
    ctx: Context, 
    text: str, 
    bold: bool = False, 
    italic: bool = False, 
    underline: bool = False,
    font_size: Optional[int] = None,
    font_name: Optional[str] = None,
    color: Optional[str] = None,
    alignment: Optional[str] = None
) -> str:
    """
    向文档添加段落文本
    
    Parameters:
    - text: 段落文本内容
    - bold: 是否加粗
    - italic: 是否斜体
    - underline: 是否下划线
    - font_size: 字体大小（磅）
    - font_name: 字体名称
    - color: 文本颜色（格式如：#FF0000）
    - alignment: 对齐方式 (left, center, right, justify)
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        # 添加段落
        paragraph = processor.current_document.add_paragraph(text)
        
        # 应用额外的格式
        if paragraph.runs:
            run = paragraph.runs[0]
            run.bold = bold
            run.italic = italic
            run.underline = underline
            
            # 设置字体大小
            if font_size:
                run.font.size = Pt(font_size)
            
            # 设置字体名称
            if font_name:
                run.font.name = font_name
                # 设置中文字体
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            
            # 设置字体颜色
            if color and color.startswith('#') and len(color) == 7:
                r = int(color[1:3], 16)
                g = int(color[3:5], 16)
                b = int(color[5:7], 16)
                run.font.color.rgb = RGBColor(r, g, b)
        
        # 设置对齐方式
        if alignment:
            if alignment == "left":
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            elif alignment == "center":
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            elif alignment == "right":
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            elif alignment == "justify":
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        
        return "段落已添加"
    except Exception as e:
        error_msg = f"添加段落失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def add_heading(ctx: Context, text: str, level: int) -> str:
    """
    向文档添加标题
    
    Parameters:
    - text: 标题文本
    - level: 标题级别（1-9）
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        processor.current_document.add_heading(text, level=level)
        
        return f"已添加级别为 {level} 的标题"
    except Exception as e:
        error_msg = f"添加标题失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def add_table(ctx: Context, rows: int, cols: int, data: Optional[list] = None) -> str:
    """
    向文档添加表格
    
    Parameters:
    - rows: 表格行数
    - cols: 表格列数
    - data: 表格数据，二维数组
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        table = processor.current_document.add_table(rows=rows, cols=cols, style="Table Grid")
        
        # 填充表格数据
        if data:
            for i, row_data in enumerate(data):
                if i < rows:
                    row = table.rows[i]
                    for j, cell_text in enumerate(row_data):
                        if j < cols:
                            row.cells[j].text = str(cell_text)
        
        return f"已添加 {rows}x{cols} 表格"
    except Exception as e:
        error_msg = f"添加表格失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def get_document_info(ctx: Context) -> str:
    """
    获取文档信息，包括段落数、表格数、样式等
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        # 获取文档基本信息
        sections_count = len(doc.sections)
        paragraphs_count = len(doc.paragraphs)
        tables_count = len(doc.tables)
        
        # 获取样式列表
        paragraph_styles = []
        for style in doc.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                paragraph_styles.append(style.name)
        
        # 构建信息字符串
        info = f"文档路径: {processor.current_file_path}\n"
        info += f"节数量: {sections_count}\n"
        info += f"段落数量: {paragraphs_count}\n"
        info += f"表格数量: {tables_count}\n"
        info += f"可用段落样式: {', '.join(paragraph_styles[:10])}..."
        
        return info
    except Exception as e:
        error_msg = f"获取文档信息失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def search_text(ctx: Context, keyword: str) -> str:
    """
    在文档中搜索文本
    
    Parameters:
    - keyword: 要搜索的关键词
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        results = []
        
        # 在段落中搜索
        for i, paragraph in enumerate(doc.paragraphs):
            if keyword in paragraph.text:
                results.append({
                    "type": "段落",
                    "index": i,
                    "text": paragraph.text
                })
        
        # 在表格中搜索
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    if keyword in cell.text:
                        results.append({
                            "type": "表格单元格",
                            "table_index": t_idx,
                            "row": r_idx,
                            "column": c_idx,
                            "text": cell.text
                        })
        
        if not results:
            return f"未找到关键词 '{keyword}'"
        
        # 构建响应
        response = f"找到 {len(results)} 处包含 '{keyword}' 的内容：\n\n"
        for idx, result in enumerate(results):
            response += f"{idx+1}. {result['type']} "
            if result['type'] == "段落":
                response += f"索引 {result['index']}: {result['text'][:100]}"
                if len(result['text']) > 100:
                    response += "..."
                response += "\n"
            else:
                response += f"在表格 {result['table_index']} 的单元格 ({result['row']},{result['column']}): {result['text'][:100]}"
                if len(result['text']) > 100:
                    response += "..."
                response += "\n"
        
        return response
    except Exception as e:
        error_msg = f"搜索文本失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def search_and_replace(ctx: Context, keyword: str, replace_with: str, preview_only: bool = False) -> str:
    """
    在文档中搜索文本并替换，提供详细的替换信息和预览选项
    
    Parameters:
    - keyword: 要搜索的关键词
    - replace_with: 替换为的文本
    - preview_only: 是否仅预览而不实际替换，默认为False
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        results = []
        
        # 在段落中搜索
        for i, paragraph in enumerate(doc.paragraphs):
            if keyword in paragraph.text:
                # 保存原始文本和替换后的文本
                original_text = paragraph.text
                replaced_text = original_text.replace(keyword, replace_with)
                results.append({
                    "type": "段落",
                    "index": i,
                    "original": original_text,
                    "replaced": replaced_text,
                    "count": original_text.count(keyword)
                })
                
                # 如果不是预览模式，则执行替换
                if not preview_only:
                    paragraph.text = replaced_text
        
        # 在表格中搜索
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    if keyword in cell.text:
                        # 保存原始文本和替换后的文本
                        original_text = cell.text
                        replaced_text = original_text.replace(keyword, replace_with)
                        results.append({
                            "type": "表格单元格",
                            "table_index": t_idx,
                            "row": r_idx,
                            "column": c_idx,
                            "original": original_text,
                            "replaced": replaced_text,
                            "count": original_text.count(keyword)
                        })
                        
                        # 如果不是预览模式，则执行替换
                        if not preview_only:
                            # 替换表格单元格中的所有段落
                            for para in cell.paragraphs:
                                if keyword in para.text:
                                    para.text = para.text.replace(keyword, replace_with)
        
        if not results:
            return f"未找到关键词 '{keyword}'"
        
        # 计算总替换次数
        total_replacements = sum(item["count"] for item in results)
        
        # 构建响应
        action_word = "预览" if preview_only else "替换"
        response = f"{action_word}将 '{keyword}' 替换为 '{replace_with}'，共找到 {len(results)} 处位置，{total_replacements} 次出现：\n\n"
        
        for idx, result in enumerate(results):
            response += f"{idx+1}. 在{result['type']} "
            if result['type'] == "段落":
                response += f"索引 {result['index']} 中{action_word} {result['count']} 次:\n"
            else:
                response += f"表格 {result['table_index']} 的单元格 ({result['row']},{result['column']}) 中{action_word} {result['count']} 次:\n"
            
            # 显示替换前后的文本片段（上下文）
            max_display = 50
            if len(result['original']) > max_display * 2:
                # 查找关键词的位置并显示其周围的文本
                start_pos = result['original'].find(keyword)
                start_pos = max(0, start_pos - max_display)
                excerpt_original = "..." + result['original'][start_pos:start_pos + max_display * 2] + "..."
                excerpt_replaced = "..." + result['replaced'][start_pos:start_pos + max_display * 2] + "..."
            else:
                excerpt_original = result['original']
                excerpt_replaced = result['replaced']
            
            response += f"  原文: {excerpt_original}\n"
            response += f"  替换后: {excerpt_replaced}\n\n"
        
        if preview_only:
            response += "这是替换预览，未实际执行替换。要执行替换，请将preview_only设置为False。"
        else:
            response += "替换已执行完成。"
        
        return response
    except Exception as e:
        error_msg = f"搜索并替换失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def find_and_replace(ctx: Context, find_text: str, replace_text: str) -> str:
    """
    在文档中查找并替换文本
    
    Parameters:
    - find_text: 要查找的文本
    - replace_text: 替换为的文本
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        replace_count = 0
        
        # 在段落中查找和替换
        for paragraph in doc.paragraphs:
            if find_text in paragraph.text:
                paragraph.text = paragraph.text.replace(find_text, replace_text)
                replace_count += paragraph.text.count(replace_text)
        
        # 在表格中查找和替换
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if find_text in paragraph.text:
                            paragraph.text = paragraph.text.replace(find_text, replace_text)
                            replace_count += paragraph.text.count(replace_text)
        
        return f"已将 '{find_text}' 替换为 '{replace_text}'，共替换 {replace_count} 处"
    except Exception as e:
        error_msg = f"查找替换失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def merge_table_cells(
    ctx: Context,
    table_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int
) -> str:
    """
    合并表格单元格
    
    Parameters:
    - table_index: 表格索引
    - start_row: 起始行索引
    - start_col: 起始列索引
    - end_row: 结束行索引
    - end_col: 结束列索引
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        if not doc.tables:
            return "文档中没有表格"
        
        if table_index < 0 or table_index >= len(doc.tables):
            return f"表格索引超出范围: {table_index}，文档共有 {len(doc.tables)} 个表格"
        
        table = doc.tables[table_index]
        
        # 检查行列索引是否有效
        if start_row < 0 or start_row >= len(table.rows):
            return f"起始行索引超出范围: {start_row}，表格共有 {len(table.rows)} 行"
        
        if start_col < 0 or start_col >= len(table.columns):
            return f"起始列索引超出范围: {start_col}，表格共有 {len(table.columns)} 列"
        
        if end_row < start_row or end_row >= len(table.rows):
            return f"结束行索引无效: {end_row}，应在 {start_row} 到 {len(table.rows)-1} 之间"
        
        if end_col < start_col or end_col >= len(table.columns):
            return f"结束列索引无效: {end_col}，应在 {start_col} 到 {len(table.columns)-1} 之间"
        
        # 获取起始单元格和结束单元格
        start_cell = table.cell(start_row, start_col)
        end_cell = table.cell(end_row, end_col)
        
        # 合并单元格
        start_cell.merge(end_cell)
        
        return f"已合并表格 {table_index} 的单元格从 ({start_row},{start_col}) 到 ({end_row},{end_col})"
    except Exception as e:
        error_msg = f"合并表格单元格失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def split_table(ctx: Context, table_index: int, row_index: int) -> str:
    """
    在指定行后拆分表格为两个表格
    
    Parameters:
    - table_index: 表格索引
    - row_index: 在此行后拆分表格
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        if not doc.tables:
            return "文档中没有表格"
        
        if table_index < 0 or table_index >= len(doc.tables):
            return f"表格索引超出范围: {table_index}，文档共有 {len(doc.tables)} 个表格"
        
        table = doc.tables[table_index]
        
        if row_index < 0 or row_index >= len(table.rows) - 1:
            return f"行索引无效: {row_index}，应在 0 到 {len(table.rows)-2} 之间"
        
        # 使用底层XML操作拆分表格
        # 获取表格元素
        tbl = table._tbl
        
        # 计算拆分位置
        split_position = row_index + 1
        
        # 创建新表格元素
        new_tbl = OxmlElement('w:tbl')
        
        # 复制表格属性
        for child in tbl.xpath('./w:tblPr')[0].getchildren():
            new_tbl.append(child.copy())
        
        # 复制表格网格设置
        for child in tbl.xpath('./w:tblGrid')[0].getchildren():
            new_tbl.append(child.copy())
        
        # 移动行到新表格
        rows = tbl.xpath('./w:tr')
        for i in range(split_position, len(rows)):
            new_tbl.append(rows[i])
        
        # 在原表格后插入新表格
        tbl.addnext(new_tbl)
        
        return f"已在行 {row_index} 后拆分表格 {table_index}"
    except Exception as e:
        error_msg = f"拆分表格失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def add_table_row(ctx: Context, table_index: int, data: Optional[list] = None) -> str:
    """
    向表格添加一行
    
    Parameters:
    - table_index: 表格索引
    - data: 行数据，列表形式
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        if not doc.tables:
            return "文档中没有表格"
        
        if table_index < 0 or table_index >= len(doc.tables):
            return f"表格索引超出范围: {table_index}，文档共有 {len(doc.tables)} 个表格"
        
        table = doc.tables[table_index]
        
        # 添加新行
        new_row = table.add_row()
        
        # 填充行数据
        if data:
            for i, cell_text in enumerate(data):
                if i < len(new_row.cells):
                    new_row.cells[i].text = str(cell_text)
        
        return f"已向表格 {table_index} 添加新行"
    except Exception as e:
        error_msg = f"添加表格行失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def delete_table_row(ctx: Context, table_index: int, row_index: int) -> str:
    """
    删除表格中的一行
    
    Parameters:
    - table_index: 表格索引
    - row_index: 要删除的行索引
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        if not doc.tables:
            return "文档中没有表格"
        
        if table_index < 0 or table_index >= len(doc.tables):
            return f"表格索引超出范围: {table_index}，文档共有 {len(doc.tables)} 个表格"
        
        table = doc.tables[table_index]
        
        if row_index < 0 or row_index >= len(table.rows):
            return f"行索引超出范围: {row_index}，表格共有 {len(table.rows)} 行"
        
        # 使用底层XML操作删除行
        row = table.rows[row_index]._tr
        row.getparent().remove(row)
        
        return f"已删除表格 {table_index} 的第 {row_index} 行"
    except Exception as e:
        error_msg = f"删除表格行失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def edit_table_cell(ctx: Context, table_index: int, row_index: int, col_index: int, text: str) -> str:
    """
    编辑表格单元格内容
    
    Parameters:
    - table_index: 表格索引
    - row_index: 行索引
    - col_index: 列索引
    - text: 单元格文本
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        if not doc.tables:
            return "文档中没有表格"
        
        if table_index < 0 or table_index >= len(doc.tables):
            return f"表格索引超出范围: {table_index}，文档共有 {len(doc.tables)} 个表格"
        
        table = doc.tables[table_index]
        
        if row_index < 0 or row_index >= len(table.rows):
            return f"行索引超出范围: {row_index}，表格共有 {len(table.rows)} 行"
        
        if col_index < 0 or col_index >= len(table.columns):
            return f"列索引超出范围: {col_index}，表格共有 {len(table.columns)} 列"
        
        # 修改单元格内容
        table.cell(row_index, col_index).text = text
        
        return f"表格 {table_index} 的单元格 ({row_index}, {col_index}) 已修改"
    except Exception as e:
        error_msg = f"编辑表格单元格失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def add_page_break(ctx: Context) -> str:
    """
    添加分页符
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        processor.current_document.add_page_break()
        
        return "已添加分页符"
    except Exception as e:
        error_msg = f"添加分页符失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def set_page_margins(
    ctx: Context,
    top: Optional[float] = None,
    bottom: Optional[float] = None,
    left: Optional[float] = None,
    right: Optional[float] = None
) -> str:
    """
    设置页面边距
    
    Parameters:
    - top: 上边距（厘米）
    - bottom: 下边距（厘米）
    - left: 左边距（厘米）
    - right: 右边距（厘米）
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        # 获取当前节（默认使用第一个节）
        section = doc.sections[0]
        
        # 设置页面边距
        if top is not None:
            section.top_margin = Cm(top)
        if bottom is not None:
            section.bottom_margin = Cm(bottom)
        if left is not None:
            section.left_margin = Cm(left)
        if right is not None:
            section.right_margin = Cm(right)
        
        return "页面边距已设置"
    except Exception as e:
        error_msg = f"设置页面边距失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def delete_paragraph(ctx: Context, paragraph_index: int) -> str:
    """
    删除文档中的指定段落
    
    Parameters:
    - paragraph_index: 要删除的段落索引
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return f"段落索引超出范围: {paragraph_index}，文档共有 {len(doc.paragraphs)} 个段落"
        
        # python-docx不提供直接删除段落的方法，需要使用底层XML操作
        paragraph = doc.paragraphs[paragraph_index]
        p = paragraph._element
        p.getparent().remove(p)
        # 删除段落对象的引用，以便垃圾回收
        paragraph._p = None
        paragraph._element = None
        
        return f"段落 {paragraph_index} 已删除"
    except Exception as e:
        error_msg = f"删除段落失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def delete_text(ctx: Context, paragraph_index: int, start_pos: int, end_pos: int) -> str:
    """
    删除文档中指定段落的部分文本
    
    Parameters:
    - paragraph_index: 段落索引
    - start_pos: 开始位置（从0开始计数）
    - end_pos: 结束位置（不包含该位置的字符）
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return f"段落索引超出范围: {paragraph_index}，文档共有 {len(doc.paragraphs)} 个段落"
        
        paragraph = doc.paragraphs[paragraph_index]
        text = paragraph.text
        
        if start_pos < 0 or start_pos >= len(text):
            return f"起始位置超出范围: {start_pos}，段落长度为 {len(text)}"
        
        if end_pos <= start_pos or end_pos > len(text):
            return f"结束位置无效: {end_pos}，应在 {start_pos+1} 到 {len(text)} 之间"
        
        # 构建新文本（删除指定部分）
        new_text = text[:start_pos] + text[end_pos:]
        paragraph.text = new_text
        
        return f"已删除段落 {paragraph_index} 中从位置 {start_pos} 到 {end_pos} 的文本"
    except Exception as e:
        error_msg = f"删除文本失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def save_as_document(ctx: Context, new_file_path: str) -> str:
    """
    将当前文档保存为新文件
    
    Parameters:
    - new_file_path: 新文件保存路径
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        # 保存为新文件
        processor.current_document.save(new_file_path)
        
        # 更新当前文件路径
        processor.current_file_path = new_file_path
        processor.documents[new_file_path] = processor.current_document
        
        return f"文档已保存为: {new_file_path}"
    except Exception as e:
        error_msg = f"保存文档失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def create_document_copy(ctx: Context, suffix: str = "-副本") -> str:
    """
    在原文件所在目录创建当前文档的副本
    
    Parameters:
    - suffix: 添加到原文件名后的后缀，默认为"-副本"
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        if not processor.current_file_path:
            return "当前文档未保存，无法创建副本"
        
        # 解析原文件路径
        file_dir = os.path.dirname(processor.current_file_path)
        file_name = os.path.basename(processor.current_file_path)
        file_name_without_ext, file_ext = os.path.splitext(file_name)
        
        # 创建新文件名
        new_file_name = f"{file_name_without_ext}{suffix}{file_ext}"
        new_file_path = os.path.join(file_dir, new_file_name)
        
        # 保存为新文件
        processor.current_document.save(new_file_path)
        
        return f"已创建文档副本: {new_file_path}"
    except Exception as e:
        error_msg = f"创建文档副本失败: {str(e)}"
        logger.error(error_msg)
        return error_msg

@mcp.tool()
def replace_section(ctx: Context, section_title: str, new_content: list, preserve_title: bool = True) -> str:
    """
    在文档中查找特定标题并替换该标题下的内容，保持原有位置、格式和样式
    
    Parameters:
    - section_title: 要查找的标题文本
    - new_content: 新内容列表，每个元素是一个段落
    - preserve_title: 是否保留原标题，默认为True
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        # 查找标题位置
        title_index = -1
        for i, paragraph in enumerate(doc.paragraphs):
            if section_title in paragraph.text:
                title_index = i
                break
        
        if title_index == -1:
            return f"未找到标题: '{section_title}'"
        
        # 确定该部分的结束位置（下一个相同或更高级别的标题）
        end_index = len(doc.paragraphs)
        title_style = doc.paragraphs[title_index].style
        
        for i in range(title_index + 1, len(doc.paragraphs)):
            # 如果找到下一个相同级别或更高级别的标题，则设置为结束位置
            if doc.paragraphs[i].style.name.startswith('Heading') and \
               (doc.paragraphs[i].style.name <= title_style.name or doc.paragraphs[i].style == title_style):
                end_index = i
                break
        
        # 保存原始段落的样式和格式信息
        original_styles = []
        for i in range(start_delete := (title_index + (1 if preserve_title else 0)), min(end_index, start_delete + len(new_content))):
            if i < len(doc.paragraphs):
                para = doc.paragraphs[i]
                style_info = {
                    'style': para.style,
                    'alignment': para.alignment,
                    'runs': []
                }
                
                # 保存每个run的格式
                for run in para.runs:
                    run_info = {
                        'bold': run.bold,
                        'italic': run.italic,
                        'underline': run.underline,
                        'font_size': run.font.size,
                        'font_name': run.font.name,
                        'color': run.font.color.rgb if run.font.color.rgb else None
                    }
                    style_info['runs'].append(run_info)
                
                original_styles.append(style_info)
            else:
                # 如果原始段落数量不足，使用最后一个段落的样式
                if original_styles:
                    original_styles.append(original_styles[-1])
                else:
                    # 如果没有原始样式，使用默认样式
                    original_styles.append({
                        'style': None,
                        'alignment': None,
                        'runs': []
                    })
        
        # 如果原始样式数量不足，使用最后一个样式填充
        while len(original_styles) < len(new_content):
            if original_styles:
                original_styles.append(original_styles[-1])
            else:
                original_styles.append({
                    'style': None,
                    'alignment': None,
                    'runs': []
                })
        
        # 记录插入位置
        insert_position = start_delete
        
        # 从后向前删除，避免索引变化
        for i in range(end_index - 1, start_delete - 1, -1):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
        
        # 添加新内容，应用原始格式
        for i, content in enumerate(reversed(new_content)):
            # 创建新段落
            p = doc.add_paragraph()
            
            # 应用原始段落样式
            style_info = original_styles[len(new_content) - i - 1]
            if style_info['style']:
                p.style = style_info['style']
            if style_info['alignment'] is not None:
                p.alignment = style_info['alignment']
            
            # 添加文本并应用格式
            if style_info['runs'] and len(style_info['runs']) > 0:
                # 如果有多个run，尝试保持格式
                # 简化处理：将整个内容添加到一个run中，应用第一个run的格式
                run = p.add_run(content)
                run_info = style_info['runs'][0]
                
                run.bold = run_info['bold']
                run.italic = run_info['italic']
                run.underline = run_info['underline']
                
                if run_info['font_size']:
                    run.font.size = run_info['font_size']
                
                if run_info['font_name']:
                    run.font.name = run_info['font_name']
                    # 设置中文字体
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), run_info['font_name'])
                
                if run_info['color']:
                    run.font.color.rgb = run_info['color']
            else:
                # 如果没有run信息，直接添加文本
                p.text = content
            
            # 移动新段落到正确位置
            doc._body._body.insert(insert_position, p._p)
            
            # 删除原来添加的段落（在文档末尾）
            doc._body._body.remove(doc.paragraphs[-1]._p)
        
        return f"已替换标题 '{section_title}' 下的内容，保留了原有格式和样式"
    except Exception as e:
        error_msg = f"替换内容失败: {str(e)}"
        logger.error(error_msg)
        traceback.print_exc()  # 打印详细错误信息
        return error_msg

@mcp.tool()
def edit_section_by_keyword(ctx: Context, keyword: str, new_content: list, section_range: int = 3) -> str:
    """
    在文档中查找包含关键词的段落，并替换该段落及其周围内容，保持原有位置、格式和样式
    
    Parameters:
    - keyword: 要查找的关键词
    - new_content: 新内容列表，每个元素是一个段落
    - section_range: 要替换的关键词周围段落范围，默认为3
    """
    try:
        if not processor.current_document:
            return "没有打开的文档"
        
        doc = processor.current_document
        
        # 查找关键词位置
        keyword_indices = []
        for i, paragraph in enumerate(doc.paragraphs):
            if keyword in paragraph.text:
                keyword_indices.append(i)
        
        if not keyword_indices:
            return f"未找到关键词: '{keyword}'"
        
        # 使用第一个匹配项
        keyword_index = keyword_indices[0]
        
        # 确定要替换的段落范围
        start_index = max(0, keyword_index - section_range)
        end_index = min(len(doc.paragraphs), keyword_index + section_range + 1)
        
        # 保存原始段落的样式和格式信息
        original_styles = []
        for i in range(start_index, min(end_index, start_index + len(new_content))):
            if i < len(doc.paragraphs):
                para = doc.paragraphs[i]
                style_info = {
                    'style': para.style,
                    'alignment': para.alignment,
                    'runs': []
                }
                
                # 保存每个run的格式
                for run in para.runs:
                    run_info = {
                        'bold': run.bold,
                        'italic': run.italic,
                        'underline': run.underline,
                        'font_size': run.font.size,
                        'font_name': run.font.name,
                        'color': run.font.color.rgb if run.font.color.rgb else None
                    }
                    style_info['runs'].append(run_info)
                
                original_styles.append(style_info)
            else:
                # 如果原始段落数量不足，使用最后一个段落的样式
                if original_styles:
                    original_styles.append(original_styles[-1])
                else:
                    # 如果没有原始样式，使用默认样式
                    original_styles.append({
                        'style': None,
                        'alignment': None,
                        'runs': []
                    })
        
        # 如果原始样式数量不足，使用最后一个样式填充
        while len(original_styles) < len(new_content):
            if original_styles:
                original_styles.append(original_styles[-1])
            else:
                original_styles.append({
                    'style': None,
                    'alignment': None,
                    'runs': []
                })
        
        # 记录插入位置
        insert_position = start_index
        
        # 从后向前删除，避免索引变化
        for i in range(end_index - 1, start_index - 1, -1):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
        
        # 添加新内容，应用原始格式
        for i, content in enumerate(reversed(new_content)):
            # 创建新段落
            p = doc.add_paragraph()
            
            # 应用原始段落样式
            style_info = original_styles[len(new_content) - i - 1]
            if style_info['style']:
                p.style = style_info['style']
            if style_info['alignment'] is not None:
                p.alignment = style_info['alignment']
            
            # 添加文本并应用格式
            if style_info['runs'] and len(style_info['runs']) > 0:
                # 如果有多个run，尝试保持格式
                # 简化处理：将整个内容添加到一个run中，应用第一个run的格式
                run = p.add_run(content)
                run_info = style_info['runs'][0]
                
                run.bold = run_info['bold']
                run.italic = run_info['italic']
                run.underline = run_info['underline']
                
                if run_info['font_size']:
                    run.font.size = run_info['font_size']
                
                if run_info['font_name']:
                    run.font.name = run_info['font_name']
                    # 设置中文字体
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), run_info['font_name'])
                
                if run_info['color']:
                    run.font.color.rgb = run_info['color']
            else:
                # 如果没有run信息，直接添加文本
                p.text = content
            
            # 移动新段落到正确位置
            doc._body._body.insert(insert_position, p._p)
            
            # 删除原来添加的段落（在文档末尾）
            doc._body._body.remove(doc.paragraphs[-1]._p)
        
        return f"已替换包含关键词 '{keyword}' 的段落及其周围内容，保留了原有格式和样式"
    except Exception as e:
        error_msg = f"替换内容失败: {str(e)}"
        logger.error(error_msg)
        traceback.print_exc()  # 打印详细错误信息
        return error_msg

# 添加更多工具...

if __name__ == "__main__":
    # 运行MCP服务器
    mcp.run() 