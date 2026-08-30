#!/usr/bin/env bash
#
# Генератор тестового .docx для сквозной проверки просмотрщика.
#
# Зачем: полноценный документ из Word тянуть в репозиторий не хочется (бинарь,
# лицензия шрифтов, размер). Здесь мы собираем минимальный, но осмысленный
# WordprocessingML вручную и упаковываем в ZIP — .docx это и есть ZIP.
#
# Что попадает в документ (намеренно покрываем узкие места вёрстки):
#   - заголовок по центру,
#   - абзац со смешанным форматированием (bold / italic / underline / цвет),
#   - длинный абзац с выключкой по ширине — проверка растягивания пробелов,
#   - абзац с висячим отступом (hanging indent),
#   - принудительный перенос строки <w:br/>,
#   - таблица 3×3 с рамками и заливкой шапки,
#   - нумерованный и маркированный списки,
#   - принудительный разрыв страницы + текст на второй странице.
#
# Использование:  bash fixtures/make-sample-docx.sh [выходной_файл]
# По умолчанию:   fixtures/sample.docx

set -euo pipefail

OUT="${1:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/sample.docx}"
BUILD="$(mktemp -d)"
trap 'rm -rf "$BUILD"' EXIT

mkdir -p "$BUILD/_rels" "$BUILD/word/_rels"

# ---------------------------------------------------------------- Content Types
cat > "$BUILD/[Content_Types].xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
XML

# ---------------------------------------------------------------- Корневые связи
cat > "$BUILD/_rels/.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
XML

cat > "$BUILD/word/_rels/document.xml.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
XML

# ---------------------------------------------------------------- Стили
# docDefaults задаёт шрифт и кегль по умолчанию — парсер обязан их подхватить
# и применить к прогонам без явного w:rPr.
cat > "$BUILD/word/styles.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <!-- Именованный стиль: проверяем цепочку наследования docDefaults -> стиль -> прямое форматирование -->
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/></w:rPr>
  </w:style>
</w:styles>
XML

# ---------------------------------------------------------------- Нумерация
cat > "$BUILD/word/numbering.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
XML

# ---------------------------------------------------------------- Тело документа
cat > "$BUILD/word/document.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>

    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Тестовый документ DOCX</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:t xml:space="preserve">Обычный текст, затем </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>полужирный</w:t></w:r>
      <w:r><w:t xml:space="preserve">, </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>курсив</w:t></w:r>
      <w:r><w:t xml:space="preserve">, </w:t></w:r>
      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>подчёркнутый</w:t></w:r>
      <w:r><w:t xml:space="preserve">, </w:t></w:r>
      <w:r><w:rPr><w:color w:val="C00000"/></w:rPr><w:t>красный</w:t></w:r>
      <w:r><w:t xml:space="preserve">, </w:t></w:r>
      <w:r><w:rPr><w:strike/></w:rPr><w:t>зачёркнутый</w:t></w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:jc w:val="both"/><w:ind w:firstLine="567"/></w:pPr>
      <w:r><w:t>Этот абзац выключен по ширине, чтобы проверить растягивание пробелов между словами. Он должен занять несколько строк, а последняя строка обязана остаться выключенной влево — ровно так поступает Microsoft Word. Если последняя строка окажется растянутой на всю ширину, значит в layout-движке ошибка. Добавим ещё немного текста, чтобы строк гарантированно получилось хотя бы три или четыре.</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:r><w:t>Абзац с висячим отступом: первая строка начинается левее остальных. Продолжение текста должно быть сдвинуто вправо относительно первой строки.</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:t>Строка до принудительного переноса.</w:t></w:r>
      <w:r><w:br/><w:t>Строка после переноса — это тот же абзац.</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Первый пункт нумерованного списка</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Второй пункт нумерованного списка</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>Пункт маркированного списка</w:t></w:r>
    </w:p>

    <w:tbl>
      <w:tblPr>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:color="333333"/>
          <w:left w:val="single" w:sz="8" w:color="333333"/>
          <w:bottom w:val="single" w:sz="8" w:color="333333"/>
          <w:right w:val="single" w:sz="8" w:color="333333"/>
          <w:insideH w:val="single" w:sz="4" w:color="999999"/>
          <w:insideV w:val="single" w:sz="4" w:color="999999"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3360"/>
      </w:tblGrid>
      <w:tr>
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc><w:tcPr><w:shd w:fill="DDEBF7"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Колонка A</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="DDEBF7"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Колонка B</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="DDEBF7"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Колонка C</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Ячейка 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Ячейка 2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Длинный текст в ячейке, который обязан перенестись на несколько строк внутри своей колонки.</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Объединённые ячейки (colSpan = 2)</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Ячейка 3</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>

    <w:p>
      <w:pPr><w:pageBreakBefore/></w:pPr>
      <w:r><w:t>Это текст на второй странице — он проверяет обработку принудительного разрыва страницы.</w:t></w:r>
    </w:p>

    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"/>
    </w:sectPr>

  </w:body>
</w:document>
XML

# ---------------------------------------------------------------- Упаковка в ZIP
# [Content_Types].xml обязан идти первым и без сжатия (-0) — этого требует
# спецификация OPC, иначе часть парсеров откажется открывать файл.
rm -f "$OUT"
( cd "$BUILD" && zip -q -X -0 "$OUT" '[Content_Types].xml' \
  && zip -q -X -9 -r "$OUT" . -x '[Content_Types].xml' )

echo "Готово: $OUT"
unzip -l "$OUT"
