const JSZip = require('jszip');

export async function docxMailMerge(templateUriString, jsonData) {
  async function getNextRelationshipId(zip, relsPath) {
    const relsContent = await zip.file(relsPath).async('string');
    const idMatches = [...relsContent.matchAll(/Id="rId(\d+)"/g)];
    const maxId = Math.max(...idMatches.map(match => parseInt(match[1], 10)), 0);
    return `rId${maxId + 1}`;
  }

  function cmToEmu(cm) {
    return Math.round(cm * 360000);
  }

  function generateTableXml(tableData) {
    const numCols = tableData[0].length;
    const totalRelativeWidth = 5000;
    const colWidths = Array(numCols).fill(0).map(() => Math.floor(totalRelativeWidth / numCols));

    let sumOfColWidths = colWidths.reduce((sum, width) => sum + width, 0);
    let diff = totalRelativeWidth - sumOfColWidths;

    if (diff !== 0) {
      colWidths[colWidths.length - 1] += diff;
    }
    const gridXml = colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('');

    const rowsXml = tableData.map(row => {
      const cellsXml = row.map((cellContent, colIndex) => {
        return `
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${colWidths[colIndex]}" w:type="dxa"/>
              <w:tcBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
              </w:tcBorders>
            </w:tcPr>
            <w:p>
              <w:pPr>
                <w:jc w:val="center"/>
              </w:pPr>
              <w:r>
                <w:t>${cellContent}</w:t>
              </w:r>
            </w:p>
          </w:tc>
        `;
      });

      return `<w:tr>${cellsXml.join('')}</w:tr>`;
    }).join('');

    return `
      <w:tbl>
        <w:tblPr>
          <w:tblW w:w="5000" w:type="pct"/>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          </w:tblBorders>
          <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
        </w:tblPr>
        <w:tblGrid>${gridXml}</w:tblGrid>
        ${rowsXml}
      </w:tbl>
    `;
  }
  
  function generateImageXml(imageRelId, cx, cy) {
    return `
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="${cx}" cy="${cy}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="${parseInt(imageRelId.replace('rId', ''), 10)}" name="Figura1"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:nvPicPr>
                  <pic:cNvPr id="${parseInt(imageRelId.replace('rId', ''), 10)}" name="Figura1"/>
                  <pic:cNvPicPr>
                    <a:picLocks noChangeAspect="1"/>
                  </pic:cNvPicPr>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="${imageRelId}"/>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </pic:blipFill>
                <pic:spPr bwMode="auto">
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="${cx}" cy="${cy}"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                  <a:noFill/>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    `;
  }

  async function processImage(zip, documentXml, placeholder, imageBase64, width, height) {
    const imageType = imageBase64.match(/data:([^;]+);/)[1];
    const imageExtension = imageType.split('/')[1];
    const imageFileName = `image-${Date.now()}.${imageExtension}`;
    const imageData = atob(imageBase64.replace(/^data:image\/\w+;base64,/, ''));
    const imageBytes = new Uint8Array(imageData.length);

    for (let i = 0; i < imageData.length; i++) {
      imageBytes[i] = imageData.charCodeAt(i);
    }

    const relsPath = 'word/_rels/document.xml.rels';
    const imageRelId = await getNextRelationshipId(zip, relsPath);
    zip.file(`word/media/${imageFileName}`, imageBytes);

    const cx = cmToEmu(width || 3);
    const cy = cmToEmu(height || 3);
    const imageXml = generateImageXml(imageRelId, cx, cy);

    documentXml = documentXml.replace(new RegExp(`<w:t>${placeholder}</w:t>`, 'g'), imageXml);

    let relsXml = await zip.file(relsPath).async('string');
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${imageRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${imageFileName}"/>` + '</Relationships>');
    zip.file(relsPath, relsXml);

    const contentTypesPath = '[Content_Types].xml';
    let contentTypesXml = await zip.file(contentTypesPath).async('string');
    const overrideEntry = `<Override PartName="/word/media/${imageFileName}" ContentType="${imageType}"/>`;
    if (!contentTypesXml.includes(overrideEntry)) {
      contentTypesXml = contentTypesXml.replace('</Types>', `${overrideEntry}</Types>`);
    }
    zip.file(contentTypesPath, contentTypesXml);

    return documentXml;
  }

  function replaceText(documentXml, placeholder, content) {
    const formatRegex = new RegExp(`(<w:pPr>.*?<\/w:pPr>)`);
    const formatMatch = documentXml.match(formatRegex);
    const formatting = formatMatch ? formatMatch[1] : '<w:pPr/>';
    let contentWithParagraphs = content.replace(
      /\n/g,
      `</w:t></w:r></w:p><w:p>${formatting}<w:r><w:t>`
    );
    return documentXml.replace(
      new RegExp(`(<w:t>.*?${placeholder}.*?<\/w:t>)`, 'g'),
      (match, p1) => {
        return p1.replace(placeholder, `${contentWithParagraphs}`);
      }
    );
  }

  function replaceTable(documentXml, placeholder, content) {
    return documentXml.replace(
      new RegExp(`<w:p>((?!<w:p>).)*?${placeholder}([\\s\\S]*?)<\/w:p>`, 'g'),
      generateTableXml(content)
    );
  }

  try {
    const base64Prefix = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,";
    if (!templateUriString.startsWith(base64Prefix)) {
      throw new Error("O templateUriString fornecido não é um base64 de um arquivo .docx válido.");
    }
    const templateBase64 = templateUriString.substring(base64Prefix.length);

    const zip = await JSZip.loadAsync(templateBase64, { base64: true });
    const documentXmlPath = 'word/document.xml';
    let documentXml = await zip.file(documentXmlPath).async('string');

    for (const item of jsonData) {
      const { type, placeholder, content, width, height } = item;

      switch (type) {
        case 'text':
          documentXml = replaceText(documentXml, placeholder, content);
          break;

        case 'image':
          documentXml = await processImage(zip, documentXml, placeholder, content, width, height);
          break;

        case 'table':
          documentXml = replaceTable(documentXml, placeholder, content);
          break;
      }
    }

    zip.file(documentXmlPath, documentXml);
    const base64String = await zip.generateAsync({ type: 'base64' });

    return `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${base64String}`;

  } catch (error) {
    console.error('Erro durante o Mail Merge:', error.message);
    throw error;
  }
}
