#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS Excel修复工具 - 一体化版本
包含所有必需的功能模块
"""

# ========== 核心修复模块 ==========
import zipfile
import os
import re
import uuid
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.utils import column_index_from_string


# ========== xlsx 包后处理工具：写入双方原生「单元格内图片」结构 ==========
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
NS_PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
IMAGE_REL_TYPE = NS_R + '/image'

# 图片扩展名 -> 包内容类型（注入 [Content_Types].xml 用）
CONTENT_TYPE_BY_EXT = {
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'gif': 'image/gif', 'bmp': 'image/bmp', 'tif': 'image/tiff',
    'tiff': 'image/tiff', 'emf': 'image/x-emf', 'wmf': 'image/x-wmf',
    'svg': 'image/svg+xml', 'webp': 'image/webp',
}

# Excel 2024 真实输出的静态部件（「置于单元格」图片的结构定义，逐字保留）
RDRICHVALUE_STRUCTURE_XML = (
    XML_DECL +
    '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" '
    'count="1"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/>'
    '<k n="CalcOrigin" t="i"/></s></rvStructures>')

RD_RICH_VALUE_TYPES_XML = (
    XML_DECL +
    '<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" '
    'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><global><keyFlags>'
    '<key name="_Self"><flag name="ExcludeFromFile" value="1"/>'
    '<flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_DisplayString"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_Flags"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_Format"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_SubLabel"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_Attribution"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_Icon"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_Display"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_CanonicalPropertyNames"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '<key name="_ClassificationId"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    '</keyFlags></global></rvTypesInfo>')


def sheet_part_map(parts):
    """{工作表名: 包内部件路径}；parts 为 {部件名: bytes}，需含 workbook.xml 及其 rels。"""
    wbx = ET.fromstring(parts['xl/workbook.xml'])
    rels = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    rid_target = {rel.get('Id'): rel.get('Target') for rel in rels}
    result = {}
    for el in wbx.iter():
        if el.tag.rsplit('}', 1)[-1] != 'sheet':
            continue
        target = rid_target.get(el.get('{' + NS_R + '}id'), '')
        if target.startswith('/'):
            part = target.lstrip('/')
        else:
            part = os.path.normpath(os.path.join('xl', target)).replace('\\', '/')
        result[el.get('name')] = part
    return result


def _ensure_content_types(xml_bytes, defaults=None, overrides=None):
    """向 [Content_Types].xml 注入缺失的 Default / Override（幂等）。"""
    text = xml_bytes.decode('utf-8')
    for ext, content_type in (defaults or {}).items():
        if 'Extension="%s"' % ext not in text:
            text = text.replace(
                '</Types>',
                '<Default Extension="%s" ContentType="%s"/></Types>' % (ext, content_type))
    for part, content_type in (overrides or {}).items():
        if 'PartName="%s"' % part not in text:
            text = text.replace(
                '</Types>',
                '<Override PartName="%s" ContentType="%s"/></Types>' % (part, content_type))
    return text.encode('utf-8')


def _add_relationships(xml_bytes, entries):
    """向 .rels 注入新关系（分配现有最大 rId 之后的编号；同 Type+Target 已存在则跳过）。"""
    text = xml_bytes.decode('utf-8')
    max_id = 0
    for m in re.finditer(r'Id="rId(\d+)"', text):
        max_id = max(max_id, int(m.group(1)))
    add = ''
    for rel_type, target in entries:
        if rel_type in text and target in text:
            continue
        max_id += 1
        add += '<Relationship Id="rId%d" Type="%s" Target="%s"/>' % (max_id, rel_type, target)
    if add:
        text = text.replace('</Relationships>', add + '</Relationships>')
    return text.encode('utf-8')


def _patch_cell_vm(xml_bytes, vm_by_coord):
    """为工作表 XML 中指定单元格的 <c> 元素追加 vm="N"（置于单元格图片引用）。"""
    text = xml_bytes.decode('utf-8')
    for coord, slot in vm_by_coord.items():
        pattern = re.compile(r'<c r="%s"([^>]*?)>' % re.escape(coord))
        m = pattern.search(text)
        if m is None:
            raise ValueError("工作表 XML 中未找到单元格 %s" % coord)
        if 'vm=' in m.group(1):
            raise ValueError("单元格 %s 已存在 vm 属性" % coord)
        if m.group(1).rstrip().endswith('/'):
            raise ValueError("单元格 %s 是自闭合空元素，无法挂 vm" % coord)
        text = pattern.sub(
            lambda mm: '<c r="%s"%s vm="%d">' % (coord, mm.group(1), slot),
            text, count=1)
    return text.encode('utf-8')


def _patch_dispimg_cell(xml_bytes, cells_by_coord):
    """把 openpyxl 写出的 DISPIMG 单元格补成 WPS 真品的存储形态：
    t="str" + 缓存值 <v>=DISPIMG(&quot;ID&quot;,1)</v>（对照真实 WPS 文件）。"""
    text = xml_bytes.decode('utf-8')
    for coord, image_id in cells_by_coord.items():
        pattern = re.compile(r'<c r="%s"([^>]*?)>(.*?)</c>' % re.escape(coord), re.S)
        m = pattern.search(text)
        if m is None:
            raise ValueError("工作表 XML 中未找到单元格 %s" % coord)
        attrs = re.sub(r'\s+t="[^"]*"', '', m.group(1))
        inner = re.sub(r'<v\s*/>|<v>.*?</v>', '', m.group(2), flags=re.S)
        cached = '=DISPIMG(&quot;%s&quot;,1)' % image_id
        text = pattern.sub(
            lambda mm: '<c r="%s" t="str"%s>%s<v>%s</v></c>'
                       % (coord, attrs, inner, cached),
            text, count=1)
    return text.encode('utf-8')


def _image_pixel_size(data):
    """从图片文件头解析像素尺寸 (w, h)，不依赖 PIL；未知格式返回 (None, None)。"""
    try:
        if data[:8] == b'\x89PNG\r\n\x1a\n' and data[12:16] == b'IHDR':
            return (int.from_bytes(data[16:20], 'big'),
                    int.from_bytes(data[20:24], 'big'))
        if data[:6] in (b'GIF87a', b'GIF89a'):
            return (int.from_bytes(data[6:8], 'little'),
                    int.from_bytes(data[8:10], 'little'))
        if data[:2] == b'BM' and len(data) >= 26:
            return (int.from_bytes(data[18:22], 'little'),
                    abs(int.from_bytes(data[22:26], 'little', signed=True)))
        if data[:2] == b'\xff\xd8':  # JPEG: scan for SOF marker
            i = 2
            while i + 9 < len(data):
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                              0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    return (int.from_bytes(data[i + 7:i + 9], 'big'),
                            int.from_bytes(data[i + 5:i + 7], 'big'))
                if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                    i += 2
                    continue
                i += 2 + int.from_bytes(data[i + 2:i + 4], 'big')
    except Exception:
        pass
    return None, None


def _write_zip_over(src_path, tmp_path, replacements, additions):
    """按 replacements 覆盖部件、additions 追加部件，重建 xlsx 包（原地替换）。"""
    with zipfile.ZipFile(src_path, 'r') as zin, \
            zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        replaced = set()
        for item in zin.infolist():
            if item.filename in replacements:
                zout.writestr(item.filename, replacements[item.filename])
                replaced.add(item.filename)
            else:
                zout.writestr(item.filename, zin.read(item.filename))
        for name, blob in additions.items():
            if name in replaced or name in replacements:
                raise ValueError("追加部件与已有部件重名: %s" % name)
            zout.writestr(name, blob)
    os.replace(tmp_path, src_path)


class PreciseSafeWPSExcelFixer:
    """精确且安全的WPS Excel修复工具，结合perfect.py的精确计算和safe.py的安全特性"""
    
    def __init__(self, xlsx_file_path):
        self.xlsx_file_path = xlsx_file_path
        self.image_list = []
        self.workbook = None
        
    def analyze_dispimg_cells(self):
        """分析所有工作表中包含DISPIMG公式的单元格"""
        try:
            self.workbook = openpyxl.load_workbook(self.xlsx_file_path, data_only=False)
        except Exception as e:
            print(f"无法加载Excel文件: {e}")
            return {}
            
        all_dispimg_cells = {}
        
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            dispimg_cells = []
            
            try:
                for row in sheet.iter_rows(min_row=1, values_only=False):
                    for cell in row:
                        if (cell.value and isinstance(cell.value, str) and 
                            '=_xlfn.DISPIMG(' in cell.value):
                            formula = cell.value
                            start = formula.find('"') + 1
                            end = formula.find('"', start)
                            image_id = formula[start:end]
                            
                            cell_info = {
                                'cell': cell,
                                'sheet': sheet,
                                'sheet_name': sheet_name,
                                'row': cell.row,
                                'column': cell.column,
                                'image_id': image_id,
                                'coordinate': cell.coordinate
                            }
                            dispimg_cells.append(cell_info)
                            self.image_list.append(image_id)
            except Exception as e:
                print(f"分析工作表 {sheet_name} 时出错: {e}")
                continue
                
            if dispimg_cells:
                all_dispimg_cells[sheet_name] = dispimg_cells
                print(f"  发现 {len(dispimg_cells)} 个DISPIMG公式")
                    
        return all_dispimg_cells
    
    def get_image_mapping(self):
        """获取图片ID到文件路径的映射关系"""
        try:
            with zipfile.ZipFile(self.xlsx_file_path, 'r') as zfile:
                required_files = ['xl/cellimages.xml', 'xl/_rels/cellimages.xml.rels']
                for req_file in required_files:
                    if req_file not in zfile.namelist():
                        print(f"缺少必要文件: {req_file}")
                        return {}
                        
                xml_content = zfile.read('xl/cellimages.xml')
                relxml_content = zfile.read('xl/_rels/cellimages.xml.rels')
        except Exception as e:
            print(f"读取图片映射时出错: {e}")
            return {}

        try:
            root = ET.fromstring(xml_content)
            name_to_embed_map = {}
            
            namespaces = {
                'etc': 'http://www.wps.cn/officeDocument/2017/etCustomData',
                'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            }
            
            for cell_image in root.findall('.//etc:cellImage', namespaces):
                name_elem = cell_image.find('.//xdr:cNvPr', namespaces)
                embed_elem = cell_image.find('.//a:blip', namespaces)
                
                if name_elem is not None and embed_elem is not None:
                    name = name_elem.attrib['name']
                    embed = embed_elem.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
                    name_to_embed_map[name] = embed

            root1 = ET.fromstring(relxml_content)
            namespaces = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            id_target_map = {
                child.attrib['Id']: child.attrib.get('Target', 'No Target Found') 
                for child in root1.findall('.//r:Relationship', namespaces=namespaces)
            }
            
            name_to_target_map = {
                name: id_target_map[embed] 
                for name, embed in name_to_embed_map.items() 
                if embed in id_target_map
            }
            
            return name_to_target_map
            
        except Exception as e:
            print(f"解析XML时出错: {e}")
            return {}
    
    def extract_image_from_xlsx(self, image_path):
        """从xlsx文件中提取指定路径的图片数据"""
        try:
            with zipfile.ZipFile(self.xlsx_file_path, 'r') as zfile:
                actual_image_path = f'xl/{image_path}'
                if actual_image_path in zfile.namelist():
                    return zfile.read(actual_image_path)
        except Exception as e:
            print(f"提取图片数据时出错: {e}")
        return None
    
    def fix_excel_file_precise_safe(self, output_path=None):
        """方向一：WPS 单元格图片(DISPIMG) -> Excel 原生「置于单元格」图片。

        图片字节原样搬运（不重新编码）；Excel 以其原生渲染规则显示：
        始终等比、完整落在单元格内、上下/左右至少一边贴合，
        调整行高列宽时按 min(格宽/图宽, 格高/图高) 自适应缩放。"""
        if output_path is None:
            output_path = self.xlsx_file_path.replace('.xlsx', '_fixed.xlsx')

        print("=== WPS 单元格图片 -> Excel「置于单元格」图片 ===")
        all_dispimg_cells = self.analyze_dispimg_cells()
        if not all_dispimg_cells:
            print("未发现需要修复的DISPIMG公式")
            return None
        total_cells = sum(len(cells) for cells in all_dispimg_cells.values())
        print(f"总共发现 {total_cells} 个DISPIMG公式需要转换")

        image_mapping = self.get_image_mapping()
        if not image_mapping:
            print("未找到图片映射关系")
            return None

        # 1) 收集每个单元格对应的原始图片字节
        jobs = []          # [(sheet_name, coordinate, image_id)]
        media_blobs = {}   # image_id -> 原始图片字节（同图多格引用时只写一份媒体）
        for sheet_name, dispimg_cells in all_dispimg_cells.items():
            for cell_info in dispimg_cells:
                coordinate = cell_info['coordinate']
                image_id = cell_info['image_id']
                if image_id not in image_mapping:
                    print(f"  [FAIL] 未找到图片映射: {coordinate} -> {image_id}")
                    continue
                image_data = self.extract_image_from_xlsx(image_mapping[image_id])
                if not image_data:
                    print(f"  [FAIL] 无法提取图片数据: {coordinate} -> {image_id}")
                    continue
                jobs.append((sheet_name, coordinate, image_id))
                media_blobs[image_id] = image_data
        if not jobs:
            print("[FAIL] 没有可转换的图片")
            return None
        print(f"收集完成: {len(jobs)} 个单元格, {len(media_blobs)} 个独立图片（原样保留）")

        # 2) openpyxl 清除 DISPIMG 公式（cellimages 部件随重建包自动移除），
        #    单元格写入 #VALUE! 错误值占位（与 Excel 置于单元格图片的存储形式一致）
        for sheet_name, coordinate, image_id in jobs:
            self.workbook[sheet_name][coordinate].value = '#VALUE!'
        for sheet in self.workbook.worksheets:
            if hasattr(sheet, '_ext_lst'):
                sheet._ext_lst = None
        self.workbook.save(output_path)
        print(f"已保存基础文件: {output_path}")

        # 3) 后处理：注入 rich value 链路 + 原始媒体 + 单元格 vm 引用
        with zipfile.ZipFile(output_path, 'r') as zfile:
            parts = {name: zfile.read(name) for name in zfile.namelist()}

        # 媒体命名避开 openpyxl 已有文件，扩展名与字节完全保留
        rel_index = {}      # image_id -> richValueRel 序号(0-based)
        media_bare = []     # 与 richValueRel 顺序一致的媒体文件名
        additions = {}
        for i, (image_id, blob) in enumerate(media_blobs.items()):
            ext = os.path.splitext(image_mapping[image_id])[1].lower().lstrip('.') or 'png'
            media_name = f'wpscellimg{i + 1}.{ext}'
            additions['xl/media/' + media_name] = blob
            rel_index[image_id] = i
            media_bare.append(media_name)
            print(f"  媒体: {image_id} -> xl/media/{media_name} ({len(blob)} 字节, 不重新编码)")

        # 每格一个 rich value 槽位（vm 序号 1 起），多格同图时共享 richValueRel/媒体
        slot_count = len(jobs)
        future_bks = ''.join(
            '<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}">'
            f'<xlrd:rvb i="{k}"/></ext></extLst></bk>' for k in range(slot_count))
        value_bks = ''.join(f'<bk><rc t="1" v="{k}"/></bk>' for k in range(slot_count))
        metadata_xml = (
            XML_DECL +
            '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">'
            '<metadataTypes count="1"><metadataType name="XLRICHVALUE" '
            'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" '
            'splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
            'coerce="1"/></metadataTypes>'
            f'<futureMetadata name="XLRICHVALUE" count="{slot_count}">{future_bks}</futureMetadata>'
            f'<valueMetadata count="{slot_count}">{value_bks}</valueMetadata></metadata>')

        rv_items = ''.join(
            f'<rv s="0"><v>{rel_index[image_id]}</v><v>5</v></rv>'
            for _, _, image_id in jobs)
        rdrichvalue_xml = (
            XML_DECL +
            '<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" '
            f'count="{slot_count}">{rv_items}</rvData>')

        rel_count = len(media_bare)
        rich_value_rel_xml = (
            XML_DECL +
            '<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/'
            f'richvaluerel" xmlns:r="{NS_R}">'
            + ''.join(f'<rel r:id="rId{i + 1}"/>' for i in range(rel_count))
            + '</richValueRels>')
        rv_rels_xml = (
            XML_DECL +
            f'<Relationships xmlns="{NS_PACKAGE_RELS}">'
            + ''.join(
                f'<Relationship Id="rId{i + 1}" Type="{IMAGE_REL_TYPE}" '
                f'Target="../media/{name}"/>'
                for i, name in enumerate(media_bare))
            + '</Relationships>')

        # [Content_Types] / workbook 关系表 / 各表 vm 属性
        exts = {os.path.splitext(name)[1].lstrip('.').lower() for name in media_bare}
        defaults = {ext: CONTENT_TYPE_BY_EXT.get(ext, 'application/octet-stream') for ext in exts}
        overrides = {
            '/xl/metadata.xml':
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
            '/xl/richData/richValueRel.xml': 'application/vnd.ms-excel.richvaluerel+xml',
            '/xl/richData/rdrichvalue.xml': 'application/vnd.ms-excel.rdrichvalue+xml',
            '/xl/richData/rdrichvaluestructure.xml':
                'application/vnd.ms-excel.rdrichvaluestructure+xml',
            '/xl/richData/rdRichValueTypes.xml': 'application/vnd.ms-excel.richvaluetypes+xml',
        }
        parts['[Content_Types].xml'] = _ensure_content_types(
            parts['[Content_Types].xml'], defaults, overrides)
        parts['xl/_rels/workbook.xml.rels'] = _add_relationships(
            parts['xl/_rels/workbook.xml.rels'], [
                (NS_R + '/sheetMetadata', 'metadata.xml'),
                ('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue',
                 'richData/rdrichvalue.xml'),
                ('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure',
                 'richData/rdrichvaluestructure.xml'),
                ('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes',
                 'richData/rdRichValueTypes.xml'),
                ('http://schemas.microsoft.com/office/2022/10/relationships/richValueRel',
                 'richData/richValueRel.xml'),
            ])

        sheet_map = sheet_part_map(parts)
        vm_by_sheet = {}
        for slot, (sheet_name, coordinate, image_id) in enumerate(jobs, start=1):
            vm_by_sheet.setdefault(sheet_name, {})[coordinate] = slot
        for sheet_name, coord_map in vm_by_sheet.items():
            part = sheet_map.get(sheet_name)
            if part is None or part not in parts:
                raise ValueError(f"找不到工作表部件: {sheet_name}")
            parts[part] = _patch_cell_vm(parts[part], coord_map)

        additions.update({
            'xl/metadata.xml': metadata_xml.encode('utf-8'),
            'xl/richData/rdrichvalue.xml': rdrichvalue_xml.encode('utf-8'),
            'xl/richData/rdrichvaluestructure.xml': RDRICHVALUE_STRUCTURE_XML.encode('utf-8'),
            'xl/richData/rdRichValueTypes.xml': RD_RICH_VALUE_TYPES_XML.encode('utf-8'),
            'xl/richData/richValueRel.xml': rich_value_rel_xml.encode('utf-8'),
            'xl/richData/_rels/richValueRel.xml.rels': rv_rels_xml.encode('utf-8'),
        })
        _write_zip_over(output_path, output_path + '.tmp', parts, additions)

        print(f"\n[COMPLETE] 转换完成！总计 {total_cells}，成功 {len(jobs)}")
        print("图片已在 Excel 中变为原生「置于单元格」图片（等比自适应、不重新编码）")
        print("注意：WPS 不显示 Excel 的置于单元格图片；如需在 WPS 中查看，"
              "请用右侧「Excel → WPS」板块把该文件转回 WPS 单元格图片")
        print(f"输出文件: {output_path}")
        return output_path

    def preview_fixes(self):
        """预览将要修复的内容"""
        print("=== 预览修复内容 ===")
        all_dispimg_cells = self.analyze_dispimg_cells()
        
        if not all_dispimg_cells:
            print("未发现需要修复的内容")
            return
            
        image_mapping = self.get_image_mapping()
        
        total = 0
        for sheet_name, cells in all_dispimg_cells.items():
            print(f"\n工作表 '{sheet_name}':")
            for cell_info in cells:
                image_id = cell_info['image_id']
                if image_id in image_mapping:
                    print(f"  {cell_info['coordinate']}: {image_id} -> {image_mapping[image_id]}")
                else:
                    print(f"  {cell_info['coordinate']}: {image_id} (图片映射缺失)")
                total += 1
        
        print(f"\n总计发现 {total} 个需要修复的图片")


# ========== 反向修复模块：Excel 置于单元格图片 -> WPS 原生单元格图片 ==========
class ReverseCellImageFixer:
    """反向修复：Excel「图片置于单元格」(Place in Cell / rich value) -> WPS 原生单元格图片。

    Excel 2024 / Microsoft 365 的「图片置于单元格」把图片存为 rich value
    （xl/metadata.xml + xl/richData/* + 单元格 vm 属性），WPS 无法显示，
    单元格只呈现 #VALUE!。本修复器按元数据链路还原每个单元格对应的图片，
    清空 rich value 引用后，写入 WPS 原生 DISPIMG 公式 + xl/cellimages.xml，
    图片字节原样搬运（不重新编码），在 WPS 中与原生「嵌入单元格图片」一致：
    图片在单元格内随单元格移动，放大量以置入时大小为上限，可随格等比缩小。"""

    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff',
                  '.emf', '.wmf', '.svg')

    def __init__(self, xlsx_file_path):
        self.xlsx_file_path = xlsx_file_path
        self._helper = PreciseSafeWPSExcelFixer(xlsx_file_path)

    # ---------- XML 工具 ----------
    @staticmethod
    def _localname(tag):
        if isinstance(tag, str) and '}' in tag:
            return tag.rsplit('}', 1)[-1]
        return tag

    def _build_vm_image_map(self, zfile, names):
        """解析元数据链路，返回 {vm 值(1起): 图片部件路径}。

        链路: 单元格 vm -> valueMetadata[K-1].rc.v -> futureMetadata[XLRICHVALUE][pos]
              -> xlrd:rvb.i -> rdrichvalue 第 i 个 <rv> 的首个 <v>（richValueRel 序号）
              -> richValueRel 第 idx 个 <rel r:id> -> rels -> xl/media 图片部件
        """
        md = ET.fromstring(zfile.read('xl/metadata.xml'))

        # futureMetadata(name) -> [rich value 序号]（按 <bk> 顺序）
        rvb_by_name = {}
        for el in md.iter():
            if self._localname(el.tag) != 'futureMetadata':
                continue
            idxs = []
            for bk in el:
                if self._localname(bk.tag) != 'bk':
                    continue
                rvb_index = None
                for child in bk.iter():
                    if self._localname(child.tag) == 'rvb':
                        try:
                            rvb_index = int(child.get('i'))
                        except (TypeError, ValueError):
                            rvb_index = None
                        break
                idxs.append(rvb_index)
            rvb_by_name[el.get('name')] = idxs

        # valueMetadata -> [futureMetadata 中的序号]
        fm_positions = []
        for el in md.iter():
            if self._localname(el.tag) != 'valueMetadata':
                continue
            for bk in el:
                if self._localname(bk.tag) != 'bk':
                    continue
                pos = None
                for child in bk:
                    if self._localname(child.tag) == 'rc':
                        try:
                            pos = int(child.get('v'))
                        except (TypeError, ValueError):
                            pos = None
                        break
                fm_positions.append(pos)
        if not fm_positions:
            return {}

        # rich value 序号 -> richValueRel 序号
        rv_root = ET.fromstring(zfile.read('xl/richData/rdrichvalue.xml'))
        rv_rel_indexes = []
        for rv in rv_root:
            if self._localname(rv.tag) != 'rv':
                continue
            rel_index = None
            for child in rv:
                if self._localname(child.tag) == 'v':
                    try:
                        rel_index = int((child.text or '').strip())
                    except (TypeError, ValueError):
                        rel_index = None
                    break
            rv_rel_indexes.append(rel_index)

        # richValueRel 序号 -> rId
        rel_root = ET.fromstring(zfile.read('xl/richData/richValueRel.xml'))
        rel_ids = []
        for rel in rel_root:
            if self._localname(rel.tag) != 'rel':
                continue
            rel_ids.append(rel.get(f'{{{self.R_NS}}}id'))

        # rId -> 部件路径
        rid_to_target = {}
        rels_part = 'xl/richData/_rels/richValueRel.xml.rels'
        if rels_part in names:
            for rel in ET.fromstring(zfile.read(rels_part)):
                if self._localname(rel.tag) == 'Relationship':
                    rid_to_target[rel.get('Id')] = rel.get('Target')

        # 组装完整链路（仅保留图片类型部件）
        rvb_list = rvb_by_name.get('XLRICHVALUE')
        if rvb_list is None and rvb_by_name:
            rvb_list = next(iter(rvb_by_name.values()))
        img_map = {}
        for vm_value, pos in enumerate(fm_positions, start=1):
            try:
                if rvb_list is None or pos is None or not (0 <= pos < len(rvb_list)):
                    continue
                rel_index = rv_rel_indexes[rvb_list[pos]]
                target = rid_to_target.get(rel_ids[rel_index])
                if not target:
                    continue
                part = os.path.normpath(
                    os.path.join('xl/richData', target)).replace('\\', '/')
                if part in names and part.lower().endswith(self.IMAGE_EXTS):
                    img_map[vm_value] = part
            except (IndexError, KeyError, TypeError, ValueError):
                continue
        return img_map

    def _sheet_name_to_part(self, zfile):
        """工作表名 -> 部件路径（xl/workbook.xml + workbook.xml.rels）"""
        rels = {}
        for rel in ET.fromstring(zfile.read('xl/_rels/workbook.xml.rels')):
            if self._localname(rel.tag) == 'Relationship':
                rels[rel.get('Id')] = rel.get('Target')
        mapping = {}
        for el in ET.fromstring(zfile.read('xl/workbook.xml')).iter():
            if self._localname(el.tag) != 'sheet':
                continue
            target = rels.get(el.get(f'{{{self.R_NS}}}id'))
            if not target:
                continue
            if target.startswith('/'):
                # openpyxl 会写绝对 Target（/xl/worksheets/...），部件名不带前导斜杠
                part = target.lstrip('/')
            else:
                part = os.path.normpath(os.path.join('xl', target)).replace('\\', '/')
            mapping[el.get('name')] = part
        return mapping

    # ---------- 分析 ----------
    def analyze_cell_image_cells(self):
        """扫描所有工作表中带 vm 属性且链路可解析为图片的单元格。
        返回 {sheet_name: [{'coordinate', 'row', 'column', 'media_part'}]}"""
        result = {}
        try:
            zfile = zipfile.ZipFile(self.xlsx_file_path, 'r')
        except Exception as e:
            print(f"无法打开Excel文件: {e}")
            return {}
        with zfile:
            names = set(zfile.namelist())
            required = ('xl/metadata.xml', 'xl/richData/rdrichvalue.xml',
                        'xl/richData/richValueRel.xml')
            if not all(n in names for n in required):
                return {}  # 包内不存在置于单元格图片（rich value）结构
            try:
                vm_map = self._build_vm_image_map(zfile, names)
            except Exception as e:
                print(f"解析富值元数据时出错: {e}")
                return {}
            if not vm_map:
                return {}
            for sheet_name, sheet_part in self._sheet_name_to_part(zfile).items():
                if sheet_part not in names:
                    continue
                try:
                    root = ET.fromstring(zfile.read(sheet_part))
                except Exception:
                    continue
                infos = []
                for el in root.iter():
                    if self._localname(el.tag) != 'c':
                        continue
                    vm = el.get('vm')
                    if not vm:
                        continue
                    coordinate = el.get('r') or ''
                    media = vm_map.get(int(vm)) if vm.isdigit() else None
                    if not media:
                        print(f"  工作表 {sheet_name} 单元格 {coordinate}: "
                              "富值引用无法解析为图片，跳过")
                        continue
                    match = re.match(r'([A-Z]+)(\d+)$', coordinate)
                    if not match:
                        continue
                    infos.append({
                        'coordinate': coordinate,
                        'column': column_index_from_string(match.group(1)),
                        'row': int(match.group(2)),
                        'sheet_name': sheet_name,
                        'media_part': media,
                    })
                if infos:
                    result[sheet_name] = infos
        return result

    # ---------- 转换 ----------
    # ---------- 转换 ----------
    def convert_excel_to_wps(self, output_path=None):
        """方向二：Excel「置于单元格」图片(rich value) -> WPS 原生单元格图片(DISPIMG)。

        图片字节原样搬运（不重新编码）；WPS 以其原生渲染规则显示：
        图片在单元格内并随单元格移动；放大量以置入时大小为上限，
        单元格缩小时按比例缩小。"""
        if output_path is None:
            base, ext = os.path.splitext(self.xlsx_file_path)
            output_path = base + '_fixed' + (ext or '.xlsx')

        print("=== Excel「置于单元格」图片 -> WPS 单元格图片 ===")
        cells = self.analyze_cell_image_cells()
        if not cells:
            print("未发现「图片置于单元格」(rich value) 图片，无需修复")
            return None
        total = sum(len(v) for v in cells.values())
        print(f"共发现 {total} 个置于单元格图片，开始转换...")

        # 1) 收集每个单元格对应的原始图片字节（保持原扩展名）
        jobs = []          # [(sheet_name, coordinate, image_id, media_part)]
        media_blobs = {}   # media_part -> 原始图片字节（多格同图时只写一份媒体）
        with zipfile.ZipFile(self.xlsx_file_path, 'r') as zfile:
            k = 0
            for sheet_name, infos in cells.items():
                for info in infos:
                    k += 1
                    # 与 WPS 真品一致: ID_ + 32 位大写十六进制
                    image_id = 'ID_' + uuid.uuid4().hex.upper()
                    media_part = info['media_part']
                    jobs.append((sheet_name, info['coordinate'], image_id, media_part))
                    if media_part not in media_blobs:
                        media_blobs[media_part] = zfile.read(media_part)
        if not jobs:
            print("[FAIL] 没有可转换的图片")
            return None
        print(f"收集完成: {len(jobs)} 个单元格, {len(media_blobs)} 个独立图片（原样保留）")

        # 2) openpyxl 写入 DISPIMG 公式（rich value 结构随重建包自动移除）
        #    真品公式为双参数 _xlfn.DISPIMG("ID_xxx",1)
        wb = openpyxl.load_workbook(self.xlsx_file_path)
        for sheet_name, coordinate, image_id, _media_part in jobs:
            wb[sheet_name][coordinate].value = f'=_xlfn.DISPIMG("{image_id}",1)'
        wb.save(output_path)
        print(f"已保存基础文件: {output_path}")

        # 3) 后处理：注入 cellimages.xml + 原始媒体
        with zipfile.ZipFile(output_path, 'r') as zfile:
            parts = {name: zfile.read(name) for name in zfile.namelist()}

        # rId 与媒体按 media_part 去重；cellImage 条目按单元格一一对应
        rid_of_media = {}
        rel_entries = []    # [(rId 序号, 媒体部件路径)]
        additions = {}
        for media_part, blob in media_blobs.items():
            idx = len(rel_entries) + 1
            rid_of_media[media_part] = idx
            ext = os.path.splitext(media_part)[1].lower().lstrip('.') or 'png'
            media_name = f'wpsimg{idx}.{ext}'
            additions['xl/media/' + media_name] = blob
            rel_entries.append((idx, media_name))
            print(f"  媒体: {media_part} -> xl/media/{media_name} ({len(blob)} 字节, 不重新编码)")

        # 与 WPS 真品逐字段一致的 cellImage 结构（nvPicPr/cNvPicPr、
        # blipFill/stretch、spPr/xfrm+prstGeom+noFill+ln），对照真实文件
        # 土山汉墓遗址展陈文物选配清单.xlsx 抓取。
        def _entry(n, image_id, media_part):
            w, h = _image_pixel_size(media_blobs[media_part])
            descr = re.sub(r'[^\w.\-]', '_',
                           os.path.basename(media_part)) or 'image'
            return (
                '<etc:cellImage><xdr:pic><xdr:nvPicPr>'
                f'<xdr:cNvPr id="{n}" name="{image_id}" descr="{descr}"/>'
                '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>'
                '</xdr:nvPicPr><xdr:blipFill>'
                f'<a:blip r:embed="rId{rid_of_media[media_part]}"/>'
                '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>'
                '<a:xfrm><a:off x="0" y="0"/>'
                f'<a:ext cx="{(w or 100) * 9525}" cy="{(h or 100) * 9525}"/>'
                '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                '<a:noFill/><a:ln w="9525"><a:noFill/></a:ln>'
                '</xdr:spPr></xdr:pic></etc:cellImage>')

        cell_entries = ''.join(
            _entry(n, image_id, media_part)
            for n, (_sheet, _coord, image_id, media_part) in enumerate(jobs, start=1))
        cellimages_xml = (
            XML_DECL +
            '<etc:cellImages '
            'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
            'spreadsheetDrawing" '
            f'xmlns:r="{NS_R}" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">'
            + cell_entries + '</etc:cellImages>')
        cellimages_rels_xml = (
            XML_DECL +
            f'<Relationships xmlns="{NS_PACKAGE_RELS}">'
            + ''.join(
                f'<Relationship Id="rId{idx}" Type="{IMAGE_REL_TYPE}" '
                f'Target="media/{name}"/>'
                for idx, name in rel_entries)
            + '</Relationships>')

        exts = {os.path.splitext(name)[1].lstrip('.').lower() for name in additions}
        defaults = {ext: CONTENT_TYPE_BY_EXT.get(ext, 'application/octet-stream') for ext in exts}
        parts['[Content_Types].xml'] = _ensure_content_types(
            parts['[Content_Types].xml'], defaults,
            {'/xl/cellimages.xml':
                'application/vnd.wps-officedocument.cellimage+xml'})

        # 真品 workbook.xml.rels 含指向 cellimages.xml 的 WPS 专属关系；
        # 缺此关系时 WPS 直接忽略该部件（实测不渲染任何图片）。
        parts['xl/_rels/workbook.xml.rels'] = _add_relationships(
            parts['xl/_rels/workbook.xml.rels'],
            [('http://www.wps.cn/officeDocument/2020/cellImage',
              'cellimages.xml')])

        # 单元格补 t="str" 与缓存值（真品形态）
        sheet_map = sheet_part_map(parts)
        cells_by_sheet = {}
        for sheet_name, coordinate, image_id, _media in jobs:
            cells_by_sheet.setdefault(sheet_name, {})[coordinate] = image_id
        for sheet_name, coord_map in cells_by_sheet.items():
            part = sheet_map.get(sheet_name)
            if part is None or part not in parts:
                raise ValueError(f"找不到工作表部件: {sheet_name}")
            parts[part] = _patch_dispimg_cell(parts[part], coord_map)

        additions.update({
            'xl/cellimages.xml': cellimages_xml.encode('utf-8'),
            'xl/_rels/cellimages.xml.rels': cellimages_rels_xml.encode('utf-8'),
        })
        _write_zip_over(output_path, output_path + '.tmp', parts, additions)

        print(f"\n[COMPLETE] 转换完成！总计 {total}，成功 {len(jobs)}")
        print("图片已在 WPS 中变为原生单元格图片（随格移动、放大约束到置入大小、缩小等比）")
        print("注意：Excel 不显示 WPS 的 DISPIMG 单元格图片；如需在 Excel 中查看，"
              "请用左侧「WPS → Excel」板块把该文件转回 Excel 置于单元格图片")
        print(f"输出文件: {output_path}")
        return output_path


def detect_direction(xlsx_path):
    """自动判别修复方向：'wps2excel'=WPS DISPIMG 文件；'excel2wps'=置于单元格图片文件；
    None=包内未发现特殊图片结构（默认按 WPS 方向处理）"""
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zfile:
            names = set(zfile.namelist())
    except Exception:
        return None
    if 'xl/cellimages.xml' in names:
        return 'wps2excel'
    if 'xl/metadata.xml' in names and 'xl/richData/richValueRel.xml' in names:
        return 'excel2wps'
    return None


# ========== GUI主程序 ==========
# 界面设计遵循 DESIGN.md（来源：github.com/VoltAgent/awesome-design-md → design-md/apple/DESIGN.md）
# 色彩 / 字体 / 圆角 / 间距均引用该规范令牌，Tk 映射与扩展说明见项目根目录 DESIGN.md
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
import os
import sys
import queue
import threading

try:
    # 拖放支持（tkinterdnd2）；不可用时自动降级为“点击选择文件”
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False


APP_TITLE = "WPS Excel 图片修复工具"
APP_VERSION = "2.0"

# ===== Apple DESIGN.md 颜色令牌（唯一强调色：Action Blue #0066cc） =====
C = {
    'primary': '#0066cc',         # colors.primary — 全部“可点击”元素
    'primary_focus': '#0071e3',   # colors.primary-focus — 悬停/聚焦
    'ink': '#1d1d1f',             # colors.ink / body
    'ink_muted_80': '#333333',    # colors.ink-muted-80
    'ink_muted_48': '#7a7a7a',    # colors.ink-muted-48 — 次级说明文字
    'hairline': '#e0e0e0',        # colors.hairline — 卡片 1px 描边
    'divider_soft': '#f0f0f0',    # colors.divider-soft — 进度条轨道
    'canvas': '#ffffff',          # colors.canvas — 窗口底色
    'parchment': '#f5f5f7',       # colors.canvas-parchment — 拖放区表面
    'pearl': '#fafafc',           # colors.surface-pearl
    'chip': '#d2d2d7',            # colors.surface-chip-translucent — 禁用态
    'on_primary': '#ffffff',      # colors.on-primary
    'error': '#d70015',           # 语义扩展（原规范未涵盖状态色，仅用于错误反馈）
}

# ===== Apple DESIGN.md 字号令牌（CSS px × 0.75 = Tk pt；权重梯 300/400/600/700，无 500） =====
FONT = {
    'tagline': ('Microsoft YaHei UI', 16, 'bold'),        # typography.tagline 21px/600
    'body': ('Microsoft YaHei UI', 13),                   # typography.body 17px/400
    'body_strong': ('Microsoft YaHei UI', 13, 'bold'),    # typography.body-strong 17px/600
    'caption': ('Microsoft YaHei UI', 11),                # typography.caption 14px/400
    'caption_strong': ('Microsoft YaHei UI', 11, 'bold'), # typography.caption-strong 14px/600
    'button': ('Microsoft YaHei UI', 13),                 # button-primary → typography.body
    'log': ('Consolas', 10),                              # 等宽日志（SF Mono → Consolas 替代）
}

# ===== Apple DESIGN.md 圆角语法（sm 紧凑工具 / lg 卡片 / pill 主行动，不混用） =====
RADIUS = {'xs': 5, 'sm': 8, 'md': 11, 'lg': 18, 'pill': 9999}
# ===== Apple DESIGN.md 间距令牌 =====
SPACE = {'xxs': 4, 'xs': 8, 'sm': 12, 'md': 17, 'lg': 24, 'xl': 32}

FILETYPES = [("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")]


def default_output_path(input_path):
    """在原位置（源文件所在文件夹）生成输出路径：xxx.xlsx -> xxx_fixed.xlsx"""
    base, ext = os.path.splitext(input_path)
    if not ext:
        ext = ".xlsx"
    return base + "_fixed" + ext.lower()


def parse_dnd_paths(data):
    """解析 tkinterdnd2 的拖放数据，支持带空格的路径（用花括号包裹）"""
    paths, token, in_braces = [], "", False
    for ch in data:
        if in_braces:
            if ch == '}':
                in_braces = False
            else:
                token += ch
        elif ch == '{':
            in_braces = True
        elif ch.isspace():
            if token:
                paths.append(token)
                token = ""
        else:
            token += ch
    if token:
        paths.append(token)
    return paths


def _round_rect(canvas, x1, y1, x2, y2, radius, **kwargs):
    """在 Canvas 上绘制圆角矩形（smooth polygon，圆角一并被平滑）"""
    r = min(radius, (x2 - x1) / 2.0, (y2 - y1) / 2.0)
    if r <= 0:
        return canvas.create_rectangle(x1, y1, x2, y2, **kwargs)
    points = [
        x1 + r, y1,
        x2 - r, y1,
        x2, y1,
        x2, y1 + r,
        x2, y2 - r,
        x2, y2,
        x2 - r, y2,
        x1 + r, y2,
        x1, y2,
        x1, y2 - r,
        x1, y1 + r,
        x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=24, **kwargs)


class AppleButton(tk.Canvas):
    """component.button-primary：Action Blue 胶囊按钮（rounded.pill），
    按下应用系统微交互 transform: scale(0.95)"""

    def __init__(self, master, text, command=None, kind='primary',
                 font=None, pad_x=22, pad_y=11, bg=C['canvas'], **kw):
        font = font or FONT['button']
        self._font = tkfont.Font(font=font)
        self._text = text
        self._command = command
        self._kind = kind
        self._pad_x = pad_x
        self._pad_y = pad_y
        self._pressed = False
        self._hover = False
        self._enabled = True
        width = max(self._font.measure(text) + 2 * pad_x, 96)
        height = self._font.metrics('linespace') + 2 * pad_y
        super().__init__(master, width=width, height=height, bg=bg,
                         highlightthickness=0, bd=0, cursor='hand2', **kw)
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<ButtonPress-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)
        self._redraw()

    # --- 状态 ---
    def set_enabled(self, enabled):
        self._enabled = bool(enabled)
        self._pressed = False
        self._hover = False
        self.configure(cursor='hand2' if self._enabled else 'arrow')
        self._redraw()

    def _colors(self):
        if not self._enabled:
            if self._kind == 'primary':
                return C['chip'], C['ink_muted_48'], ''
            return C['canvas'], C['chip'], C['chip']
        if self._kind == 'primary':
            bg = C['primary_focus'] if (self._hover or self._pressed) else C['primary']
            return bg, C['on_primary'], ''
        # button-secondary-pill：幽灵胶囊
        bg = C['parchment'] if self._pressed else C['canvas']
        return bg, C['primary'], C['primary']

    def _redraw(self):
        self.delete('all')
        w = int(self['width'])
        h = int(self['height'])
        scale = 0.95 if self._pressed else 1.0
        dw, dh = w * scale, h * scale
        x1, y1 = (w - dw) / 2.0, (h - dh) / 2.0
        x2, y2 = x1 + dw, y1 + dh
        bg, fg, outline = self._colors()
        _round_rect(self, x1, y1, x2, y2, dh / 2.0,
                    fill=bg, outline=outline, width=1)
        self.create_text(w / 2.0, h / 2.0, text=self._text,
                         font=self._font, fill=fg)

    def _on_enter(self, event):
        if not self._enabled:
            return
        self._hover = True
        self._redraw()

    def _on_leave(self, event):
        if not self._enabled:
            return
        self._hover = False
        self._pressed = False
        self._redraw()

    def _on_press(self, event):
        if not self._enabled:
            return
        self._pressed = True
        self._redraw()

    def _on_release(self, event):
        if not self._enabled:
            return
        inside = (0 <= event.x < int(self['width'])
                  and 0 <= event.y < int(self['height']))
        was_pressed = self._pressed
        self._pressed = False
        self._redraw()
        if was_pressed and inside and self._command:
            self._command()


class LinkLabel(tk.Label):
    """component.text-link：透明底 + Action Blue 文本操作（无边框）"""

    def __init__(self, master, text, command=None, bg=C['canvas'], **kw):
        super().__init__(master, text=text, font=FONT['caption'],
                         fg=C['primary'], bg=bg,
                         activeforeground=C['primary_focus'],
                         cursor='hand2', **kw)
        self._command = command
        self.bind('<Button-1>', self._on_click)
        self.bind('<Enter>', lambda e: self.configure(fg=C['primary_focus']))
        self.bind('<Leave>', lambda e: self.configure(fg=C['primary']))

    def _on_click(self, event):
        if self._command:
            self._command()


class AppleCheckbox(tk.Frame):
    """复选框：rounded.xs(5) 方块，选中态 primary 填充 + 白色对勾，
    未选中态 1px chip 描边（hover 升级为 primary）"""

    SIZE = 16

    def __init__(self, master, text, variable, command=None, bg=C['canvas']):
        super().__init__(master, bg=bg)
        self._var = variable
        self._command = command
        self._hover = False
        self._box = tk.Canvas(self, width=self.SIZE, height=self.SIZE, bg=bg,
                              highlightthickness=0, bd=0, cursor='hand2')
        self._box.pack(side='left')
        self._label = tk.Label(self, text=text, font=FONT['caption'],
                               fg=C['ink'], bg=bg, cursor='hand2')
        self._label.pack(side='left', padx=(SPACE['xs'], 0))
        for widget in (self._box, self._label):
            widget.bind('<Button-1>', self._on_click)
            widget.bind('<Enter>', self._on_enter)
            widget.bind('<Leave>', self._on_leave)
        self._box.bind('<Configure>', lambda e: self._redraw())
        self._redraw()

    def _redraw(self):
        c = self._box
        c.delete('all')
        s, r = self.SIZE, RADIUS['xs']
        if bool(self._var.get()):
            _round_rect(c, 0, 0, s - 1, s - 1, r,
                        fill=C['primary'], outline='')
            c.create_line(3.5, 8.5, 6.5, 11.5, 12.5, 5.0,
                          fill=C['on_primary'], width=2,
                          capstyle='round', joinstyle='round')
        else:
            hovered = self._hover
            _round_rect(c, 0, 0, s - 1, s - 1, r, fill=C['canvas'],
                        outline=C['primary'] if hovered else C['chip'],
                        width=1.5 if hovered else 1)

    def _on_click(self, event=None):
        self._var.set(not self._var.get())
        self._redraw()
        if self._command:
            self._command()

    def _on_enter(self, event=None):
        self._hover = True
        self._redraw()

    def _on_leave(self, event=None):
        self._hover = False
        self._redraw()


class RoundedCard(tk.Canvas):
    """component.store-utility-card：白底 + 1px hairline + rounded.lg(18px) + 24px 内边距；
    parchment 变体（拖放区）以表面色区分，不加描边（色彩变化即分隔）"""

    def __init__(self, master, fill=None, outline=None, radius=None,
                 pad=SPACE['lg'], height=None, bg=C['canvas']):
        fill = C['canvas'] if fill is None else fill
        outline = C['hairline'] if outline is None else outline
        radius = RADIUS['lg'] if radius is None else radius
        super().__init__(master, bg=bg, highlightthickness=0, bd=0,
                         height=height or 1, width=1)
        self._fill = fill
        self._base_outline = outline
        self._outline = outline
        self._border_w = 1
        self._radius = radius
        self._pad = pad
        self.content = tk.Frame(self, bg=fill)
        self._window = self.create_window(0, 0, window=self.content, anchor='nw')
        self.bind('<Configure>', self._redraw)

    def set_active_border(self, active):
        """拖放悬停：描边升级为 2px primary-focus（chip-selected 规则）"""
        if self._base_outline == '' and not active:
            self._outline = ''
            self._border_w = 1
        elif active:
            self._outline = C['primary_focus']
            self._border_w = 2
        else:
            self._outline = self._base_outline
            self._border_w = 1
        self._redraw()

    def _redraw(self, event=None):
        w, h = self.winfo_width(), self.winfo_height()
        if w < 6 or h < 6:
            return
        self.delete('card_bg')
        _round_rect(self, 1, 1, w - 1, h - 1, self._radius,
                    fill=self._fill, outline=self._outline,
                    width=self._border_w, tags='card_bg')
        self.tag_lower('card_bg')
        self.coords(self._window, self._pad, self._pad)
        self.itemconfigure(self._window,
                           width=max(w - 2 * self._pad, 1),
                           height=max(h - 2 * self._pad, 1))


class AppleProgress(tk.Canvas):
    """进度条：rounded.pill 胶囊，轨道 divider-soft，填充 primary（唯一强调色）"""

    def __init__(self, master, height=8, bg=C['canvas']):
        super().__init__(master, height=height, bg=bg,
                         highlightthickness=0, bd=0)
        self._value = 0.0
        self.bind('<Configure>', self._redraw)

    def set_value(self, value):
        self._value = max(0.0, min(100.0, float(value)))
        self._redraw()

    def _redraw(self, event=None):
        w, h = self.winfo_width(), max(self.winfo_height(), 8)
        if w < 6:
            return
        self.delete('all')
        _round_rect(self, 0, 0, w - 1, h - 1, h / 2.0,
                    fill=C['divider_soft'], outline='')
        fill_w = w * self._value / 100.0
        if fill_w > 1:
            _round_rect(self, 0, 0, max(fill_w, h) - 1, h - 1, h / 2.0,
                        fill=C['primary'], outline='')


class QueueWriter:
    """把修复器的 print 输出导入 GUI 日志区"""

    def __init__(self, log_queue):
        self.log_queue = log_queue

    def write(self, text):
        for line in str(text).splitlines():
            line = line.strip()
            if line:
                self.log_queue.put(line)

    def flush(self):
        pass


_STDOUT_LOCK = threading.Lock()


def _repair_wps_to_excel(file_path, output_path):
    """方向一：WPS 的 DISPIMG 图片 -> Excel 原生置于单元格图片"""
    return PreciseSafeWPSExcelFixer(file_path).fix_excel_file_precise_safe(output_path)


def _repair_excel_to_wps(file_path, output_path):
    """方向二：Excel 置于单元格图片 -> WPS 原生单元格图片（DISPIMG）"""
    return ReverseCellImageFixer(file_path).convert_excel_to_wps(output_path)


class RepairPanel(tk.Frame):
    """单个修复板块：标题 + 拖放区 + 文件列表 + 处理按钮 + 进度 + 日志。
    左右两块并列，各自独立运行（DESIGN.md：同一白画布上的两个并列分区，
    分隔靠留白而非新颜色）"""

    def __init__(self, master, title, description, repair_func):
        super().__init__(master, bg=C['canvas'])
        self.repair_func = repair_func
        self.log_queue = queue.Queue()
        # 进度/完成回调同样经队列回到主线程，worker 线程不做任何跨线程 Tk 调用
        self.ui_queue = queue.Queue()
        self.running = False
        self.last_output = None

        # ----- 板块标题：body_strong + caption 说明 -----
        head = tk.Frame(self, bg=C['canvas'])
        head.pack(fill='x', pady=(0, SPACE['sm']))
        tk.Label(head, text=title, font=FONT['body_strong'],
                 fg=C['ink'], bg=C['canvas']).pack(anchor='w')
        tk.Label(head, text=description, font=FONT['caption'],
                 fg=C['ink_muted_48'], bg=C['canvas'],
                 anchor='w').pack(anchor='w', pady=(SPACE['xxs'], 0))

        # ----- 拖放区：parchment 表面 + rounded.lg -----
        self.drop_card = RoundedCard(self, fill=C['parchment'], outline='',
                                     pad=SPACE['sm'], height=80, bg=C['canvas'])
        self.drop_card.pack(fill='x')
        sub = tk.Frame(self.drop_card.content, bg=C['parchment'])
        drop_main = tk.Label(sub, text="把 .xlsx 文件拖到这里",
                             font=FONT['body_strong'], fg=C['ink'],
                             bg=C['parchment'], cursor='hand2')
        drop_main.pack()
        drop_hint = tk.Label(sub, text="或点击此区域选择文件（支持多选）",
                             font=FONT['caption'], fg=C['ink_muted_48'],
                             bg=C['parchment'], cursor='hand2')
        drop_hint.pack(pady=(SPACE['xxs'], 0))
        sub.place(relx=0.5, rely=0.5, anchor='center')
        for widget in (self.drop_card, self.drop_card.content, sub,
                       drop_main, drop_hint):
            widget.bind('<Button-1>', lambda e: self.choose_files())

        # 拖放注册目标：整个板块区域 + 拖放区各级控件
        self.dnd_targets = [self, self.drop_card, self.drop_card.content,
                            sub, drop_main, drop_hint]

        # ----- 文件卡片：store-utility-card（白 + hairline + rounded.lg + 24px 内边距） -----
        self.files_card = RoundedCard(self, bg=C['canvas'])
        self.files_card.pack(fill='both', expand=True, pady=(SPACE['md'], 0))
        fc = self.files_card.content

        header = tk.Frame(fc, bg=C['canvas'])
        header.pack(fill='x')
        self.count_label = tk.Label(header, text="待修复文件 · 0 个",
                                    font=FONT['caption_strong'],
                                    fg=C['ink'], bg=C['canvas'])
        self.count_label.pack(side='left')
        LinkLabel(header, "清空列表", self.clear_files).pack(
            side='right', padx=SPACE['sm'])
        LinkLabel(header, "移除选中", self.remove_selected).pack(
            side='right', padx=SPACE['sm'])
        LinkLabel(header, "添加文件", self.choose_files).pack(
            side='right', padx=(SPACE['sm'], 0))

        list_frame = tk.Frame(fc, bg=C['canvas'])
        list_frame.pack(fill='both', expand=True, pady=(SPACE['sm'], 0))
        self.file_list = tk.Listbox(
            list_frame, font=FONT['caption'], fg=C['ink'], bg=C['canvas'],
            selectbackground=C['primary'], selectforeground=C['on_primary'],
            activestyle='none', exportselection=False,
            selectmode='extended', height=4, relief='flat', bd=0,
            highlightthickness=0)
        list_scroll = ttk.Scrollbar(list_frame, orient='vertical',
                                    style='Apple.Vertical.TScrollbar',
                                    command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=list_scroll.set)
        self.file_list.pack(side='left', fill='both', expand=True)
        list_scroll.pack(side='right', fill='y')
        self.file_list.bind('<Double-Button-1>', lambda e: self.remove_selected())

        # ----- 控制行：主行动胶囊 + 复选框 -----
        control = tk.Frame(self, bg=C['canvas'])
        control.pack(fill='x', pady=(SPACE['md'], 0))
        self.start_button = AppleButton(control, "开始修复",
                                        command=self.start_repair)
        self.start_button.pack(side='left')
        self.auto_open_var = tk.BooleanVar(value=True)
        AppleCheckbox(control, "修复完成后自动打开文件",
                      self.auto_open_var).pack(side='right')

        # ----- 进度：percent 右对齐 + 胶囊进度条 -----
        progress_row = tk.Frame(self, bg=C['canvas'])
        progress_row.pack(fill='x', pady=(SPACE['md'], SPACE['xs']))
        self.progress_percent = tk.Label(progress_row, text="0%",
                                         font=FONT['caption_strong'],
                                         fg=C['ink'], bg=C['canvas'])
        self.progress_percent.pack(side='right')
        self.progress_bar = AppleProgress(self)
        self.progress_bar.pack(fill='x')
        self.status_label = tk.Label(self, text="就绪 · 请先添加要处理的文件",
                                     font=FONT['caption'],
                                     fg=C['ink_muted_48'], bg=C['canvas'],
                                     anchor='w')
        self.status_label.pack(fill='x', pady=(0, SPACE['md']))

        # ----- 日志卡片 -----
        self.log_card = RoundedCard(self, bg=C['canvas'])
        self.log_card.pack(fill='both', expand=True)
        lc = self.log_card.content

        log_header = tk.Frame(lc, bg=C['canvas'])
        log_header.pack(fill='x')
        tk.Label(log_header, text="处理日志", font=FONT['caption_strong'],
                 fg=C['ink'], bg=C['canvas']).pack(side='left')
        LinkLabel(log_header, "清空日志", self.clear_log).pack(side='right')

        log_body = tk.Frame(lc, bg=C['canvas'])
        log_body.pack(fill='both', expand=True, pady=(SPACE['sm'], 0))
        self.log_text = tk.Text(log_body, font=FONT['log'], fg=C['ink'],
                                bg=C['canvas'], wrap='word', height=5,
                                relief='flat', bd=0, highlightthickness=0,
                                state='disabled', padx=2, pady=2)
        log_scroll = ttk.Scrollbar(log_body, orient='vertical',
                                   style='Apple.Vertical.TScrollbar',
                                   command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll.pack(side='right', fill='y')
        self.log_text.tag_configure('error', foreground=C['error'])
        self.log_text.tag_configure('info', foreground=C['ink'])
        self.log_text.tag_configure('muted', foreground=C['ink_muted_48'])

        self.after(100, self._poll_log)

    # ---------- 拖放 ----------
    def _highlight_drop(self, active):
        try:
            self.drop_card.set_active_border(bool(active))
        except Exception:
            pass

    def _on_drop(self, event):
        self._highlight_drop(False)
        for path in parse_dnd_paths(event.data):
            self._add_file(path)

    # ---------- 文件管理 ----------
    def choose_files(self):
        if self.running:
            return
        picked = filedialog.askopenfilenames(title="选择要修复的Excel文件",
                                             filetypes=FILETYPES)
        for path in picked:
            self._add_file(path)

    def _add_file(self, path):
        path = str(path).strip().strip('"')
        if not path:
            return
        if os.path.isdir(path):
            self.append_log(f"跳过文件夹：{path}", 'error')
            return
        if not os.path.exists(path):
            self.append_log(f"文件不存在：{path}", 'error')
            return
        if not path.lower().endswith('.xlsx'):
            self.append_log(f"仅支持 .xlsx 文件，已跳过：{os.path.basename(path)}", 'error')
            return
        path = os.path.abspath(path)
        existing = [self.file_list.get(i) for i in range(self.file_list.size())]
        if path in existing:
            self.append_log(f"已在列表中：{os.path.basename(path)}", 'muted')
            return
        self.file_list.insert('end', path)
        self.file_list.see('end')
        self._update_count()

    def remove_selected(self):
        if self.running:
            return
        for index in reversed(self.file_list.curselection()):
            self.file_list.delete(index)
        self._update_count()

    def clear_files(self):
        if self.running:
            return
        self.file_list.delete(0, 'end')
        self._update_count()

    def _update_count(self):
        self.count_label.config(text=f"待修复文件 · {self.file_list.size()} 个")

    # ---------- 日志与进度 ----------
    def append_log(self, message, tag='info'):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', message + "\n", tag)
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def clear_log(self):
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')

    def _poll_log(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                # “失败 0 个”是成功总结，不应标红；真正的失败/错误才用 error 色
                has_failure = ('[FAIL]' in message or '出错' in message
                               or ('失败' in message and '失败 0' not in message))
                if has_failure:
                    tag = 'error'
                elif set(message) <= {'=', '-'} and message:
                    tag = 'muted'
                else:
                    tag = 'info'
                self.append_log(message, tag)
        except queue.Empty:
            pass
        try:
            while True:
                action = self.ui_queue.get_nowait()
                if action[0] == 'progress':
                    self._set_progress(action[1], action[2])
                elif action[0] == 'finish':
                    self._finish(action[1])
        except queue.Empty:
            pass
        self.after(100, self._poll_log)

    def _set_progress(self, value, status=None):
        value = max(0, min(100, int(value)))
        self.progress_bar.set_value(value)
        self.progress_percent.config(text=f"{value}%")
        if status:
            self.status_label.config(text=status)

    # ---------- 处理流程 ----------
    def start_repair(self):
        if self.running:
            return
        files = [self.file_list.get(i) for i in range(self.file_list.size())]
        if not files:
            messagebox.showinfo(APP_TITLE, "请先拖放或选择要处理的 .xlsx 文件")
            return
        missing = [f for f in files if not os.path.exists(f)]
        if missing:
            messagebox.showerror(APP_TITLE,
                                 "以下文件已不存在：\n" + "\n".join(missing))
            return

        self.running = True
        self.start_button.set_enabled(False)
        self.append_log("=" * 52, 'muted')
        self.append_log(f"开始处理 {len(files)} 个文件 · 结果保存在源文件所在文件夹")

        thread = threading.Thread(
            target=self._repair_worker,
            args=(files, bool(self.auto_open_var.get())), daemon=True)
        thread.start()

    def _repair_worker(self, files, auto_open):
        total = len(files)
        success, failed = 0, 0

        # 两块并行启动时对 stdout 重定向加锁，避免日志互相串块
        acquired = _STDOUT_LOCK.acquire(blocking=False)
        if not acquired:
            self.log_queue.put("另一板块正在处理，等待其完成…")
            _STDOUT_LOCK.acquire()
        old_stdout = sys.stdout
        try:
            sys.stdout = QueueWriter(self.log_queue)
            for index, file_path in enumerate(files, start=1):
                name = os.path.basename(file_path)
                output_path = default_output_path(file_path)
                self.ui_queue.put(
                    ('progress', (index - 1) / total * 100,
                     f"正在处理 ({index}/{total}) · {name}"))

                self.log_queue.put(f"【{index}/{total}】{name}")
                self.log_queue.put(f"输出文件：{output_path}")

                try:
                    result = self.repair_func(file_path, output_path)
                    if result and os.path.exists(result):
                        success += 1
                        self.last_output = result
                        self.log_queue.put(f"[OK] 修复完成：{result}")
                        if auto_open:
                            try:
                                os.startfile(result)
                            except Exception as e:
                                self.log_queue.put(f"自动打开失败：{e}")
                    else:
                        failed += 1
                        self.log_queue.put(f"[FAIL] 未能生成修复文件：{name}")
                except Exception as e:
                    failed += 1
                    self.log_queue.put(f"[FAIL] 处理 {name} 时出错：{e}")

                self.ui_queue.put(
                    ('progress', index / total * 100,
                     f"已完成 {index}/{total}"))
        finally:
            sys.stdout = old_stdout
            _STDOUT_LOCK.release()
            summary = f"全部完成 · 成功 {success} 个，失败 {failed} 个"
            self.log_queue.put("\n" + summary)
            self.ui_queue.put(('finish', summary))

    def _finish(self, summary):
        self.running = False
        self.start_button.set_enabled(True)
        self._set_progress(100, summary)


class RepairApp:
    """交互式修复界面：左右并列双板块
    左 = WPS->Excel（DISPIMG 图片），右 = Excel->WPS（置于单元格图片）
    界面规范：DESIGN.md（Apple 设计系统）"""

    def __init__(self, initial_files=None, autostart=False):
        try:
            if HAS_DND:
                self.window = TkinterDnD.Tk()
                self.dnd_ready = True
            else:
                self.window = tk.Tk()
                self.dnd_ready = False
        except Exception:
            self.window = tk.Tk()
            self.dnd_ready = False

        self.window.title(APP_TITLE)
        self.window.geometry("1160x760")
        self.window.minsize(1000, 700)
        self.window.configure(bg=C['canvas'])

        self._setup_icon()
        self._setup_styles()
        self._build_widgets()
        self._center_window()

        if initial_files:
            for f in initial_files:
                self._route_file(f)

        # --autostart：界面加载完成后自动点击各板块的“开始修复”
        if autostart:
            self.window.after(800, self._autostart)

    # ---------- 界面构建 ----------
    def _setup_icon(self):
        try:
            candidates = []
            if hasattr(sys, '_MEIPASS'):
                candidates.append(os.path.join(sys._MEIPASS, 'assset', 'icon.ico'))
            candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assset', 'icon.ico'))
            candidates.append(os.path.join(os.getcwd(), 'assset', 'icon.ico'))
            for icon_path in candidates:
                if os.path.exists(icon_path):
                    self.window.iconbitmap(icon_path)
                    break
        except Exception:
            pass

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('Apple.Vertical.TScrollbar',
                        background=C['chip'], troughcolor=C['canvas'],
                        bordercolor=C['canvas'], relief='flat',
                        borderwidth=0, arrowcolor=C['ink_muted_48'])
        style.map('Apple.Vertical.TScrollbar',
                  background=[('active', C['hairline'])],
                  arrowcolor=[('active', C['ink'])])

    def _register_dnd(self, widget, panel):
        if not self.dnd_ready:
            return False
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', panel._on_drop)
            widget.dnd_bind('<<DropEnter>>', lambda e: panel._highlight_drop(True))
            widget.dnd_bind('<<DropLeave>>', lambda e: panel._highlight_drop(False))
            return True
        except Exception:
            return False

    def _build_widgets(self):
        outer = tk.Frame(self.window, bg=C['canvas'])
        outer.pack(fill='both', expand=True, padx=SPACE['xl'], pady=SPACE['lg'])

        # ----- 页眉：tagline 21px/600 -----
        tk.Label(outer, text=APP_TITLE, font=FONT['tagline'],
                 fg=C['ink'], bg=C['canvas']).pack(anchor='w')

        # ----- 双板块容器（先建板块，DnD 注册后再按序 pack） -----
        panels = tk.Frame(outer, bg=C['canvas'])
        self.panel_left = RepairPanel(
            panels, "WPS → Excel 修复",
            "WPS 的 DISPIMG 图片 → Excel 原生置于单元格图片",
            _repair_wps_to_excel)
        self.panel_left.pack(side='left', fill='both', expand=True)
        self.panel_right = RepairPanel(
            panels, "Excel → WPS 修复",
            "Excel 置于单元格图片 → WPS 原生单元格图片",
            _repair_excel_to_wps)
        self.panel_right.pack(side='left', fill='both', expand=True,
                              padx=(SPACE['xl'], 0))

        # 拖放注册：各板块区域（按落点分流到对应板块）
        for panel in (self.panel_left, self.panel_right):
            for widget in panel.dnd_targets:
                self._register_dnd(widget, panel)
        # 根窗口兜底：落在页眉等非板块区域时，按左右位置分流
        if self.dnd_ready:
            try:
                self.window.drop_target_register(DND_FILES)
                self.window.dnd_bind('<<Drop>>', self._on_root_drop)
                self.window.dnd_bind('<<DropEnter>>', self._on_root_enter)
                self.window.dnd_bind('<<DropLeave>>', self._on_root_leave)
            except Exception:
                self.dnd_ready = False

        if self.dnd_ready:
            hint_text = "文件拖放到左右对应板块，或点击板块内区域选择"
        else:
            hint_text = "当前未加载拖放组件，请点击各板块内区域选择文件"
        tk.Label(outer,
                 text=f"双向转换 · WPS/Excel 各自原生单元格图片互转（图片不重编码） · {hint_text}",
                 font=FONT['caption'], fg=C['ink_muted_48'], bg=C['canvas'],
                 anchor='w').pack(anchor='w', pady=(SPACE['xxs'], SPACE['sm']))

        panels.pack(fill='both', expand=True)

    def _center_window(self):
        """窗口居中于主显示器（多显示器环境下 vroot 居中会偏出主屏）"""
        self.window.update_idletasks()
        w, h = 1160, 760
        sw = sh = 0
        try:
            import ctypes
            # Tk 与 Python 进程同为 DPI-unaware，GetSystemMetrics 返回一致的虚拟化单位
            sw = ctypes.windll.user32.GetSystemMetrics(0)
            sh = ctypes.windll.user32.GetSystemMetrics(1)
        except Exception:
            sw = sh = 0
        if not sw or not sh:
            sw = self.window.winfo_screenwidth()
            sh = self.window.winfo_screenheight()
        x = max((sw - w) // 2, 0)
        y = max((sh - h - 40) // 2, 0)
        self.window.geometry(f"{w}x{h}+{x}+{y}")

    # ---------- 根窗口拖放分流 ----------
    def _panel_at(self, x_root):
        """按落点横坐标判断目标板块（窗口中线为界）"""
        mid = self.window.winfo_rootx() + self.window.winfo_width() // 2
        return self.panel_left if x_root < mid else self.panel_right

    def _on_root_drop(self, event):
        panel = self._panel_at(event.x_root)
        self.panel_left._highlight_drop(False)
        self.panel_right._highlight_drop(False)
        for path in parse_dnd_paths(event.data):
            panel._add_file(path)

    def _on_root_enter(self, event):
        panel = self._panel_at(event.x_root)
        if panel is not None:
            panel._highlight_drop(True)

    def _on_root_leave(self, event):
        self.panel_left._highlight_drop(False)
        self.panel_right._highlight_drop(False)

    # ---------- 文件路由 ----------
    def _route_file(self, path):
        """按包内容自动判别方向：置于单元格图片 -> 右板块，其余 -> 左板块"""
        direction = None
        try:
            direction = detect_direction(path)
        except Exception:
            pass
        panel = self.panel_right if direction == 'excel2wps' else self.panel_left
        panel._add_file(path)

    def _autostart(self):
        for panel in (self.panel_left, self.panel_right):
            if panel.file_list.size() > 0:
                panel.start_repair()

    def run(self):
        self.window.mainloop()


def run_batch(files):
    """无界面批处理模式：--batch 文件1.xlsx 文件2.xlsx ...
    按包内容自动判别修复方向（DISPIMG -> Excel；置于单元格图片 -> WPS）"""
    total = len(files)
    success = 0
    for index, file_path in enumerate(files, start=1):
        name = os.path.basename(file_path)
        output_path = default_output_path(file_path)
        direction = detect_direction(file_path)
        tag = "Excel->WPS" if direction == 'excel2wps' else "WPS->Excel"
        print(f"\n【{index}/{total}】[{tag}] {name} -> {output_path}")
        try:
            if direction == 'excel2wps':
                fixer = ReverseCellImageFixer(file_path)
                result = fixer.convert_excel_to_wps(output_path)
            else:
                fixer = PreciseSafeWPSExcelFixer(file_path)
                result = fixer.fix_excel_file_precise_safe(output_path)
            if result and os.path.exists(result):
                success += 1
                print(f"[OK] 修复完成：{result}")
            else:
                print(f"[FAIL] 未能生成修复文件：{name}")
        except Exception as e:
            print(f"[FAIL] 处理 {name} 时出错：{e}")
    print(f"\n批处理完成：成功 {success}/{total}")
    return 0 if success == total else 1


def main():
    """主函数：支持拖拽到程序图标、--batch 无界面批处理（自动判别方向）、
    --autostart 自动开始（对已加载文件的板块生效）"""
    args = sys.argv[1:]
    batch_mode = '--batch' in args
    autostart = '--autostart' in args
    args = [a for a in args if a not in ('--batch', '--autostart')]

    files = []
    for arg in args:
        path = os.path.abspath(arg)
        if os.path.exists(path) and path.lower().endswith('.xlsx'):
            files.append(path)

    if batch_mode:
        if not files:
            print("用法: WPS_Excel_Repair_Tool.exe --batch 文件1.xlsx 文件2.xlsx ...")
            return 2
        return run_batch(files)

    app = RepairApp(initial_files=files, autostart=autostart)
    app.run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
