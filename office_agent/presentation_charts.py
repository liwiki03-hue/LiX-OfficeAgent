from decimal import Decimal
from io import BytesIO
import posixpath
import re
import stat
import zipfile
from xml.etree.ElementTree import fromstring

from .presentation_drafts import PresentationError, FONT


def chart_values(rows):
    if not 2<=len(rows)<=6 or any(len(row)!=2 for row in rows):raise PresentationError('图表需两列：分类、数值；含表头，最多5项数据')
    if any(not value.strip() or len(value)>40 or value.lstrip().startswith(('=','+','-','@','{')) for value in [*rows[0],*[row[0] for row in rows[1:]]]):raise PresentationError('分类和表头应为1—40字普通文字，不接收公式')
    if len({row[0] for row in rows[1:]})!=len(rows)-1:raise PresentationError('分类重复，请明确每个数据项')
    values=[]
    for row in rows[1:]:
        if not re.fullmatch(r'-?\d{1,9}(?:\.\d{1,6})?',row[1]):raise PresentationError('数值需明确十进制数；空白、百分号、公式、混合单位不能自动转换')
        value=Decimal(row[1])
        if Decimal(str(float(value)))!=value:raise PresentationError('数值精度无法可靠保留，请先明确舍入后的数据')
        values.append(value)
    return values


def unit_text(unit):
    return '单位：'+unit+' · 数据由用户核对，事实未独立核验'


def chart_scale(values):
    return float(min(Decimal(0),min(values)*Decimal('1.1'))),float(max(Decimal(1),max(values)*Decimal('1.1')))


def add_chart(page,item):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_DATA_LABEL_POSITION
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    values=chart_values(item.rows)
    if not item.chart_unit.strip() or not item.chart_reviewed:raise PresentationError('请填写单位并明确核对本页图表数据；这不代表事实验证')
    if any(len(row[0])>14 for row in item.rows[1:]):raise PresentationError('图表分类标签过长，请缩短至14字以内并在讲稿保留全称')
    data=CategoryChartData(number_format='0.######');data.categories=[row[0] for row in item.rows[1:]]
    data.add_series(item.rows[0][1],[float(value) for value in values])
    shape=page.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,Inches(.85),Inches(1.85),Inches(11.6),Inches(4.35),data)
    shape.name='chart-'+item.id;chart=shape.chart;chart.has_legend=False;chart.chart_style=10
    for axis in [chart.category_axis,chart.value_axis]:
        axis.tick_labels.font.name=FONT;axis.tick_labels.font.size=Pt(18)
        axis.tick_labels.font.color.rgb=RGBColor.from_string('162334')
    for axis,title in [(chart.category_axis,item.rows[0][0]),(chart.value_axis,item.rows[0][1])]:
        axis.has_title=True;axis.axis_title.text_frame.text=title
        for paragraph in axis.axis_title.text_frame.paragraphs:paragraph.font.name=FONT;paragraph.font.size=Pt(18)
    chart.value_axis.minimum_scale,chart.value_axis.maximum_scale=chart_scale(values)
    chart.value_axis.tick_labels.number_format='0.######'
    plot=chart.plots[0];plot.has_data_labels=True;labels=plot.data_labels
    labels.position=XL_DATA_LABEL_POSITION.OUTSIDE_END;labels.number_format='0.######';labels.number_format_is_linked=False
    labels.font.name=FONT;labels.font.size=Pt(18);labels.font.color.rgb=RGBColor.from_string('162334')
    return shape


def checked_zip(raw,maximum=8*1024*1024):
    archive=zipfile.ZipFile(BytesIO(raw));parts=archive.infolist();names=archive.namelist()
    if len(names)!=len(set(names)) or len(names)>500 or sum(row.file_size for row in parts)>maximum:
        archive.close();raise PresentationError('包结构或展开容量无效')
    try:
        for row in parts:
            name=row.filename
            if row.flag_bits&1 or stat.S_ISLNK(row.external_attr>>16) or name.startswith(('/','\\')) or '\\' in name or ':' in name or '..' in name.split('/') or any(token in name.lower() for token in ('vba','activex','oleobject','externallink','connections','customui')):
                raise PresentationError('不支持的演示稿部件')
            if name.endswith(('.xml','.rels')):
                data=archive.read(name);markup=data.replace(b'\x00',b'').lower()
                if b'<!doctype' in markup or b'<!entity' in markup:raise PresentationError('不允许XML实体')
                node=fromstring(data)
                if any(element.get('TargetMode','').lower()=='external' for element in node.iter()):raise PresentationError('不允许外部关系')
                if any(element.tag.rsplit('}',1)[-1] in {'instrText','fldSimple','altChunk','oleObject','OLEObject'} for element in node.iter()):raise PresentationError('不允许活动内容')
                if any(element.tag.rsplit('}',1)[-1]=='f' for element in node.iter()) and not re.fullmatch(r'ppt/charts/chart\d+\.xml',name):raise PresentationError('内嵌工作簿不能含公式')
            if name.endswith('.rels'):
                base=posixpath.dirname(posixpath.dirname(name))
                for rel in node:
                    target=rel.get('Target','')
                    if any(token in rel.get('Type','').lower() for token in ('oleobject','vbaproject','control','externallink')):raise PresentationError('不允许活动关系')
                    resolved=posixpath.normpath(posixpath.join(base,target))
                    if not target or any(char in target for char in ':\\?#') or target.startswith('/') or resolved.startswith('../') or resolved not in names:raise PresentationError('内部关系目标无效')
        return archive
    except Exception:
        archive.close();raise


def validate_chart(chart,item):
    from pptx.enum.chart import XL_CHART_TYPE
    from openpyxl import load_workbook
    expected=chart_values(item.rows)
    if not item.chart_reviewed or not item.chart_unit.strip():raise PresentationError('图表数据及单位尚未核对')
    if chart.chart_type!=XL_CHART_TYPE.COLUMN_CLUSTERED or len(chart.plots)!=1 or len(chart.series)!=1:raise PresentationError('图表类型或系列数量不符')
    if [row.label for row in chart.plots[0].categories]!=[row[0] for row in item.rows[1:]] or chart.series[0].name!=item.rows[0][1] or [Decimal(str(value)) for value in chart.series[0].values]!=expected:raise PresentationError('图表缓存与批准数据不符')
    if any(not axis.has_title or axis.axis_title.text_frame.text!=title for axis,title in [(chart.category_axis,item.rows[0][0]),(chart.value_axis,item.rows[0][1])]):raise PresentationError('图表坐标含义与批准数据不符')
    if (chart.value_axis.minimum_scale,chart.value_axis.maximum_scale)!=chart_scale(expected):raise PresentationError('图表坐标范围与批准数据不符')
    rels=list(chart.part.rels.values())
    if len(rels)!=1 or not rels[0].reltype.endswith('/package') or rels[0].is_external:raise PresentationError('图表工作簿关系不符')
    workbook_part=rels[0].target_part;name=str(workbook_part.partname).lstrip('/')
    if not re.fullmatch(r'ppt/embeddings/Microsoft_Excel_Sheet\d+\.xlsx',name):raise PresentationError('图表仅接收生成器创建的工作簿')
    raw=workbook_part.blob
    if len(raw)>65536:raise PresentationError('图表工作簿超过预算')
    allowed={'[Content_Types].xml','_rels/.rels','xl/_rels/workbook.xml.rels','xl/worksheets/sheet1.xml','xl/workbook.xml','xl/sharedStrings.xml','xl/styles.xml','xl/theme/theme1.xml','docProps/core.xml','docProps/app.xml'}
    with checked_zip(raw,256*1024) as archive:
        if set(archive.namelist())!=allowed:raise PresentationError('工作簿含未批准部件')
    book=load_workbook(BytesIO(raw),read_only=True,data_only=False,keep_links=False)
    try:
        if len(book.worksheets)!=1 or book.defined_names:raise PresentationError('工作簿结构不符')
        sheet=book.worksheets[0]
        if sheet.title!='Sheet1' or sheet.sheet_state!='visible' or sheet.max_row!=len(item.rows) or sheet.max_column!=2:raise PresentationError('工作簿范围不符')
        rows=list(sheet.iter_rows())
        if any(cell.data_type=='f' for row in rows for cell in row):raise PresentationError('工作簿不能含公式')
        if rows[0][0].value is not None or rows[0][1].value!=item.rows[0][1]:raise PresentationError('工作簿表头不符')
        for index,row in enumerate(rows[1:]):
            if row[0].value!=item.rows[index+1][0] or row[1].data_type!='n' or Decimal(str(row[1].value))!=expected[index]:raise PresentationError('工作簿数据与批准值不符')
    finally:book.close()
    c='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
    formulas={node.text for node in chart.part._element.iter(c+'f')}
    count=len(item.rows)
    if formulas!={f'Sheet1!$A$2:$A${count}','Sheet1!$B$1',f'Sheet1!$B$2:$B${count}'}:raise PresentationError('图表引用范围不符')
    external=chart.part._element.find(c+'externalData')
    update=external.find(c+'autoUpdate') if external is not None else None
    if update is None or update.get('val')!='0':raise PresentationError('图表自动更新状态不符')
    return name,str(chart.part.partname).lstrip('/')
