"""
Office Pro - CLI 入口

命令行接口，提供文档生成、模板管理等命令
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Optional

try:
    import click
    CLICK_AVAILABLE = True
except ImportError:
    CLICK_AVAILABLE = False

# 导入处理器
sys.path.insert(0, str(Path(__file__).parent))
from word_processor import WordProcessor
from excel_processor import ExcelProcessor


# ==================== CLI 定义 ====================

if CLICK_AVAILABLE:
    @click.group()
    @click.version_option(version="1.0.0", prog_name="office-pro")
    def cli():
        """Office Pro - 企业级文档自动化工具"""
        pass

    # ==================== Word 命令 ====================

    @cli.group()
    def word():
        """Word 文档操作"""
        pass

    @word.command('generate')
    @click.option('--template', '-t', required=True, help='模板文件名')
    @click.option('--data', '-d', required=True, help='数据 JSON 文件路径')
    @click.option('--output', '-o', required=True, help='输出文件路径')
    @click.option('--template-dir', help='模板目录')
    def word_generate(template, data, output, template_dir):
        """使用模板生成 Word 文档"""
        try:
            # 加载数据
            with open(data, 'r', encoding='utf-8') as f:
                context = json.load(f)
            
            # 创建处理器
            wp = WordProcessor(template_dir=template_dir)
            
            # 加载模板并渲染
            wp.load_template(template)
            wp.render_and_save(context, output)
            
            click.echo(f"✓ 文档已生成: {output}")
            
        except Exception as e:
            click.echo(f"✗ 错误: {e}", err=True)
            sys.exit(1)

    @word.command('create')
    @click.option('--output', '-o', required=True, help='输出文件路径')
    @click.option('--title', help='文档标题')
    def word_create(output, title):
        """创建空白 Word 文档"""
        try:
            wp = WordProcessor()
            wp.create_document()
            
            if title:
                wp.add_heading(title, level=1)
            
            wp.save(output)
            click.echo(f"✓ 文档已创建: {output}")
            
        except Exception as e:
            click.echo(f"✗ 错误: {e}", err=True)
            sys.exit(1)

    # ==================== Excel 命令 ====================

    @cli.group()
    def excel():
        """Excel 表格操作"""
        pass

    @excel.command('generate')
    @click.option('--template', '-t', required=True, help='模板文件名')
    @click.option('--data', '-d', required=True, help='数据 JSON 文件路径')
    @click.option('--output', '-o', required=True, help='输出文件路径')
    @click.option('--template-dir', help='模板目录')
    def excel_generate(template, data, output, template_dir):
        """使用模板生成 Excel 报表"""
        try:
            # 加载数据
            with open(data, 'r', encoding='utf-8') as f:
                context = json.load(f)
            
            # 创建处理器
            ep = ExcelProcessor(template_dir=template_dir)
            
            # 加载模板并渲染
            ep.load_template(template)
            ep.render_and_save(context, output)
            
            click.echo(f"✓ 报表已生成: {output}")
            
        except Exception as e:
            click.echo(f"✗ 错误: {e}", err=True)
            sys.exit(1)

    @excel.command('create')
    @click.option('--output', '-o', required=True, help='输出文件路径')
    @click.option('--sheets', '-s', default=1, help='工作表数量')
    def excel_create(output, sheets):
        """创建空白 Excel 工作簿"""
        try:
            ep = ExcelProcessor()
            ep.create_workbook()
            
            # 重命名第一个工作表
            if sheets > 0:
                ws = ep.get_sheet()
                ws.title = "Sheet1"
            
            # 添加更多工作表
            for i in range(2, sheets + 1):
                ep.create_sheet(f"Sheet{i}")
            
            ep.save(output)
            click.echo(f"✓ 工作簿已创建: {output}")
            
        except Exception as e:
            click.echo(f"✗ 错误: {e}", err=True)
            sys.exit(1)

    @excel.command('csv-import')
    @click.option('--input', '-i', 'input_file', required=True, help='CSV 文件路径')
    @click.option('--output', '-o', required=True, help='输出 Excel 路径')
    @click.option('--delimiter', '-d', default=',', help='分隔符')
    def csv_import(input_file, output, delimiter):
        """从 CSV 导入数据"""
        try:
            ep = ExcelProcessor()
            ep.create_workbook()
            ep.import_csv(input_file, delimiter=delimiter)
            ep.save(output)
            click.echo(f"✓ 数据已导入: {output}")
            
        except Exception as e:
            click.echo(f"✗ 错误: {e}", err=True)
            sys.exit(1)

    @excel.command('csv-export')
    @click.option('--input', '-i', 'input_file', required=True, help='Excel 文件路径')
    @click.option('--output', '-o', required=True, help='输出 CSV 路径')
    @click.option('--sheet', '-s', help='工作表名称')
    @click.option('--delimiter', '-d', default=',', help='分隔符')
    def csv_export(input_file, output, sheet, delimiter):
        """导出数据到 CSV"""
        try:
            ep = ExcelProcessor()
            ep.load_workbook(input_file)
            ep.export_csv(output, sheet=sheet, delimiter=delimiter)
            click.echo(f"✓ 数据已导出: {output}")
            
        except Exception as e:
            click.echo(f"✗ 错误: {e}", err=True)
            sys.exit(1)

    # ==================== 模板管理 ====================

    @cli.group()
    def templates():
        """模板管理"""
        pass

    @templates.command('list')
    @click.option('--type', 'template_type', type=click.Choice(['word', 'excel', 'all']), default='all')
    @click.option('--template-dir', help='模板目录')
    def list_templates(template_type, template_dir):
        """列出可用模板"""
        try:
            import os
            
            if template_dir:
                base_dir = Path(template_dir)
            else:
                # 使用默认路径
                skill_root = Path(__file__).parent.parent
                base_dir = skill_root / "assets" / "templates"
            
            templates = {'word': [], 'excel': []}
            
            if template_type in ('word', 'all'):
                word_dir = base_dir / "word"
                if word_dir.exists():
                    templates['word'] = [f.name for f in word_dir.glob("*.docx")]
            
            if template_type in ('excel', 'all'):
                excel_dir = base_dir / "excel"
                if excel_dir.exists():
                    templates['excel'] = [f.name for f in excel_dir.glob("*.xlsx")]
            
            # 输出
            if template_type in ('word', 'all'):
                click.echo("\n📄 Word 模板:")
                if templates['word']:
                    for t in sorted(templates['word']):
                        click.echo(f"  • {t}")
                else:
                    click.echo("  (无)")
            
            if template_type in ('excel', 'all'):
                click.echo("\n📊 Excel 模板:")
                if templates['excel']:
                    for t in sorted(templates['excel']):
                        click.echo(f"  • {t}")
                else:
                    click.echo("  (无)")
            
            click.echo()
            
        except Exception as e:
            click.echo(f"✗ 错误: {e}", err=True)
            sys.exit(1)

    # 启动入口
    def main():
        """CLI 入口点"""
        if not CLICK_AVAILABLE:
            print("错误: click 模块未安装。请运行: pip install click")
            sys.exit(1)
        
        cli()

else:
    # 没有 click 时的占位符
    def main():
        print("错误: click 模块未安装。请运行: pip install click")
        sys.exit(1)


if __name__ == '__main__':
    main()
