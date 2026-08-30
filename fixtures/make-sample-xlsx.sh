#!/usr/bin/env bash
#
# Генератор тестового .xlsx для сквозной проверки просмотрщика таблиц.
#
# Зачем: тащить в репозиторий книгу из Excel не хочется (бинарь, размер,
# лицензия шрифтов). Здесь мы собираем минимальный, но осмысленный
# SpreadsheetML вручную и упаковываем в ZIP — .xlsx это и есть ZIP.
#
# Что попадает в книгу (намеренно покрываем узкие места отрисовки):
#   - числа, текст, логические значения и ошибка вычисления,
#   - формула с закэшированным значением (вычислять её просмотрщик не умеет),
#   - числовые форматы: разряды, проценты, валюта, дата, время,
#   - даты по обе стороны от несуществующего 29.02.1900,
#   - объединённые ячейки,
#   - колонки разной ширины и скрытая колонка,
#   - строка увеличенной высоты и скрытая строка,
#   - закреплённая шапка (одна строка) и первая колонка,
#   - перелив длинного текста в пустые соседние ячейки,
#   - число, не помещающееся в узкую колонку (показывается решётками),
#   - перенос по словам внутри ячейки,
#   - заливки, цветной и жирный шрифт, рамки,
#   - условное форматирование: сравнение, цветовая шкала, гистограмма
#     и правило с формулой (оно намеренно не поддержано),
#   - автофильтр,
#   - картинка и столбчатая диаграмма,
#   - второй лист и скрытый третий лист.
#
# Использование:  bash fixtures/make-sample-xlsx.sh [выходной_файл]
# По умолчанию:   fixtures/sample.xlsx

set -euo pipefail

OUT="${1:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/sample.xlsx}"
BUILD="$(mktemp -d)"
trap 'rm -rf "$BUILD"' EXIT

mkdir -p "$BUILD/_rels" "$BUILD/xl/_rels" "$BUILD/xl/worksheets" \
         "$BUILD/xl/worksheets/_rels" "$BUILD/xl/drawings/_rels" \
         "$BUILD/xl/charts" "$BUILD/xl/media"

# ---------------------------------------------------------------- Content Types
cat > "$BUILD/[Content_Types].xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>
XML

# ------------------------------------------------------------ корневые связи
cat > "$BUILD/_rels/.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
XML

# ------------------------------------------------------------------- книга
# Третий лист скрыт: просмотрщик обязан показать его вкладку приглушённой.
cat > "$BUILD/xl/workbook.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <sheets>
    <sheet name="Отчёт" sheetId="1" r:id="rId1"/>
    <sheet name="Форматы" sheetId="2" r:id="rId2"/>
    <sheet name="Черновик" sheetId="3" state="hidden" r:id="rId3"/>
  </sheets>
</workbook>
XML

cat > "$BUILD/xl/_rels/workbook.xml.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
XML

# ---------------------------------------------------------------- общие строки
cat > "$BUILD/xl/sharedStrings.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="16" uniqueCount="16">
  <si><t>Квартальный отчёт «Ромашка &amp; Co»</t></si>
  <si><t>Товар</t></si>
  <si><t>Кол-во</t></si>
  <si><t>Цена</t></si>
  <si><t>Сумма</t></si>
  <si><t>Доля</t></si>
  <si><t>Болты М8</t></si>
  <si><t>Гайки М8</t></si>
  <si><t>Шайбы</t></si>
  <si><t>Итого</t></si>
  <si><t>Очень длинное название, которое не помещается в колонку и переливается вправо</t></si>
  <si><t>Перенос по словам внутри одной ячейки</t></si>
  <si><t>Скрытая колонка →</t></si>
  <si><t>Примечание</t></si>
  <si><t>Значение</t></si>
  <si><t>Формат</t></si>
</sst>
XML

# --------------------------------------------------------------------- стили
# numFmt 164 — денежный с разделителем разрядов, 165 — дата, 166 — время.
cat > "$BUILD/xl/styles.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="#,##0.00&quot; ₽&quot;"/>
    <numFmt numFmtId="165" formatCode="dd.mm.yyyy"/>
    <numFmt numFmtId="166" formatCode="hh:mm:ss"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><color rgb="FF808080"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/></border>
    <border><left style="thin"><color rgb="FF8EA9DB"/></left><right style="thin"><color rgb="FF8EA9DB"/></right><top style="thin"><color rgb="FF8EA9DB"/></top><bottom style="thin"><color rgb="FF8EA9DB"/></bottom></border>
    <border><left/><right/><top style="double"><color rgb="FF1F3864"/></top><bottom/></border>
  </borders>
  <cellXfs count="13">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1"/>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1"/>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="1"/>
    <xf numFmtId="164" fontId="4" fillId="3" borderId="2"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment wrapText="1" vertical="top"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <dxfs count="3">
    <dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
    <dxf><font><b/><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>
    <dxf><font><i/><color rgb="FF9C6500"/></font></dxf>
  </dxfs>
</styleSheet>
XML

# --------------------------------------------------------------------- тема
mkdir -p "$BUILD/xl/theme"
cat > "$BUILD/xl/theme/theme1.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>
XML

# ------------------------------------------------------------------- лист 1
# Закреплены шапка (2 строки) и первая колонка: xSplit=1, ySplit=2.
# Колонка D скрыта, строка 9 скрыта, строка 11 высокая (перенос по словам).
cat > "$BUILD/xl/worksheets/sheet1.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:G12"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43"/>
  <cols>
    <col min="1" max="1" width="22" customWidth="1"/>
    <col min="2" max="2" width="9" customWidth="1"/>
    <col min="3" max="3" width="14" customWidth="1"/>
    <col min="4" max="4" width="0" hidden="1" customWidth="1"/>
    <col min="5" max="5" width="16" customWidth="1"/>
    <col min="6" max="6" width="6" customWidth="1"/>
    <col min="7" max="7" width="30" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="28" customHeight="1">
      <c r="A1" s="1" t="s"><v>0</v></c>
    </row>
    <row r="2" ht="30" customHeight="1">
      <c r="A2" s="2" t="s"><v>1</v></c>
      <c r="B2" s="2" t="s"><v>2</v></c>
      <c r="C2" s="2" t="s"><v>3</v></c>
      <c r="D2" s="2" t="s"><v>13</v></c>
      <c r="E2" s="2" t="s"><v>4</v></c>
      <c r="F2" s="12" t="s"><v>5</v></c>
      <c r="G2" s="2" t="s"><v>13</v></c>
    </row>
    <row r="3">
      <c r="A3" s="3" t="s"><v>6</v></c>
      <c r="B3" s="4"><v>1200</v></c>
      <c r="C3" s="5"><v>12.5</v></c>
      <c r="D3" s="3" t="s"><v>12</v></c>
      <c r="E3" s="5"><f>B3*C3</f><v>15000</v></c>
      <c r="F3" s="6"><v>0.5172</v></c>
      <c r="G3" s="0" t="s"><v>10</v></c>
    </row>
    <row r="4">
      <c r="A4" s="3" t="s"><v>7</v></c>
      <c r="B4" s="4"><v>800</v></c>
      <c r="C4" s="5"><v>9.9</v></c>
      <c r="E4" s="5"><f>B4*C4</f><v>7920</v></c>
      <c r="F4" s="6"><v>0.2731</v></c>
    </row>
    <row r="5">
      <c r="A5" s="3" t="s"><v>8</v></c>
      <c r="B5" s="4"><v>15000</v></c>
      <c r="C5" s="5"><v>0.4</v></c>
      <c r="E5" s="5"><f>B5*C5</f><v>6000</v></c>
      <c r="F5" s="6"><v>0.2069</v></c>
    </row>
    <row r="6">
      <c r="A6" s="3" t="s"><v>9</v></c>
      <c r="B6" s="4"><v>17000</v></c>
      <c r="E6" s="7"><f>SUM(E3:E5)</f><v>28920</v></c>
      <c r="F6" s="6"><v>1</v></c>
    </row>
    <row r="8">
      <c r="A8" s="10" t="s"><v>13</v></c>
      <c r="B8" t="b"><v>1</v></c>
      <c r="C8" t="e"><v>#DIV/0!</v></c>
    </row>
    <row r="9" hidden="1">
      <c r="A9" t="s"><v>13</v></c>
      <c r="B9"><v>999</v></c>
    </row>
    <row r="11" ht="46" customHeight="1">
      <c r="A11" s="11" t="s"><v>11</v></c>
      <c r="C11" s="4"><v>123456789012</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:G1"/>
  </mergeCells>
  <autoFilter ref="A2:G6"/>
  <drawing r:id="rId1"/>
  <conditionalFormatting sqref="F3:F5">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FFFFFFFF"/><color rgb="FF63BE7B"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="E3:E5">
    <cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan">
      <formula>10000</formula>
    </cfRule>
    <cfRule type="cellIs" dxfId="0" priority="3" operator="lessThan">
      <formula>7000</formula>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="B3:B5">
    <cfRule type="dataBar" priority="4">
      <dataBar>
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="A3:A5">
    <cfRule type="expression" dxfId="2" priority="5">
      <formula>MOD(ROW(),2)=0</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>
XML

# ------------------------------------------------------------------- лист 2
# Числовые форматы и даты по обе стороны от несуществующего 29.02.1900.
cat > "$BUILD/xl/worksheets/sheet2.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C10"/>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="26" customWidth="1"/>
    <col min="2" max="2" width="18" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" s="12" t="s"><v>15</v></c>
      <c r="B1" s="12" t="s"><v>14</v></c>
    </row>
    <row r="2"><c r="A2" t="s"><v>15</v></c><c r="B2" s="4"><v>1234567</v></c></row>
    <row r="3"><c r="A3" t="s"><v>15</v></c><c r="B3" s="5"><v>1234.5</v></c></row>
    <row r="4"><c r="A4" t="s"><v>15</v></c><c r="B4" s="6"><v>0.1234</v></c></row>
    <row r="5"><c r="A5" t="s"><v>15</v></c><c r="B5" s="8"><v>45000</v></c></row>
    <row r="6"><c r="A6" t="s"><v>15</v></c><c r="B6" s="9"><v>0.60416666</v></c></row>
    <row r="7"><c r="A7" t="s"><v>15</v></c><c r="B7" s="8"><v>59</v></c></row>
    <row r="8"><c r="A8" t="s"><v>15</v></c><c r="B8" s="8"><v>61</v></c></row>
    <row r="9"><c r="A9" t="s"><v>15</v></c><c r="B9" s="4"><v>-4200</v></c></row>
  </sheetData>
</worksheet>
XML

# ------------------------------------------------------------------- лист 3
cat > "$BUILD/xl/worksheets/sheet3.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A1"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>13</v></c></row>
  </sheetData>
</worksheet>
XML

# ------------------------------------------------- рисунок, картинка, чарт
cat > "$BUILD/xl/worksheets/_rels/sheet1.xml.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
XML

# Картинка привязана к двум ячейкам (тянется с колонками), диаграмма — к одной
# с явным размером: покрываем оба вида якорей.
cat > "$BUILD/xl/drawings/drawing1.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>13</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1" name="Логотип"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>13</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="3200400" cy="1905000"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Диаграмма"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <a:graphic><a:graphicData><c:chart r:id="rId2"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>
XML

cat > "$BUILD/xl/drawings/_rels/drawing1.xml.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
XML

# Значения диаграммы берутся из кэша: просмотрщик формулы не вычисляет.
cat > "$BUILD/xl/charts/chart1.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Выручка по товарам</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:f>Отчёт!$E$2</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Сумма</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Отчёт!$A$3:$A$5</c:f><c:strCache><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>Болты</c:v></c:pt><c:pt idx="1"><c:v>Гайки</c:v></c:pt><c:pt idx="2"><c:v>Шайбы</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Отчёт!$E$3:$E$5</c:f><c:numCache><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>15000</c:v></c:pt><c:pt idx="1"><c:v>7920</c:v></c:pt><c:pt idx="2"><c:v>6000</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
    <c:legend/>
  </c:chart>
</c:chartSpace>
XML

# Маленький PNG 16x16 сплошного цвета: тащить бинарь в репозиторий незачем,
# проще собрать его из base64 на месте.
base64 -d > "$BUILD/xl/media/image1.png" <<'B64'
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAFklEQVR4nGN4YLCBJMQwqmFUw/DVAABbQsAQEhOr1wAAAABJRU5ErkJggg==
B64

# --------------------------------------------------------------- упаковка
rm -f "$OUT"
( cd "$BUILD" && zip -q -r -X "$OUT" '[Content_Types].xml' _rels xl )

echo "Собрано: $OUT"
