// Import required libraries
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const docx = require('docx');
const sharp = require('sharp');

// Import classes from docx
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableCell,
  TableRow,
  WidthType,
  AlignmentType,
  BorderStyle,
  ShadingType,
} = docx;

// Configuration
const TESTRAIL_URL = 'https://testrail.laya.ie';
const SESSION_TOKEN = 'bab34c7f-9464-4626-ac4c-15b0c6b19e73';
const RUN_ID = 4408;
const OUTPUT_DOC = `TestRun_${RUN_ID}_Export.docx`;
const ATTACHMENTS_FOLDER = './attachments';
const MAX_IMAGE_WIDTH = 600;

// Ensure attachments folder exists
if (!fs.existsSync(ATTACHMENTS_FOLDER)) fs.mkdirSync(ATTACHMENTS_FOLDER);

// Axios instance with session cookie
const axiosInstance = axios.create({
  baseURL: TESTRAIL_URL,
  headers: { 'Cookie': `tr_session=${SESSION_TOKEN}` },
});

// Mapping TestRail status codes to labels and colors
const status_map = {
  1: { label: "Passed", color: "00FF00" },
  2: { label: "Blocked", color: "FFA500" },
  3: { label: "Untested", color: "D3D3D3" },
  4: { label: "Retest", color: "FFFF00" },
  5: { label: "Failed", color: "FF0000" },
};

// Create document with default metadata
const doc = new Document({
  creator: 'TestRail Exporter',
  title: `Test Run ${RUN_ID} Export`,
  description: 'Exported test run from TestRail',
  sections: [],
});

// Fetch all test details and build document
fetchTests();

async function fetchTests() {
  try {
    // Fetch test run metadata
    const runRes = await axiosInstance.get(`/index.php?/api/v2/get_run/${RUN_ID}`);
    const run = runRes.data;

    const testRes = await axiosInstance.get(`/index.php?/api/v2/get_tests/${RUN_ID}`);
    const tests = testRes.data.tests;

    // Group tests by status
    const grouped = { Passed: [], Blocked: [], Untested: [], Retest: [], Failed: [], Unknown: [] };
    tests.forEach(test => {
      const statusLabel = status_map[test.status_id]?.label || 'Unknown';
      grouped[statusLabel].push(test);
    });

    const totalTests = tests.length;
    const passedTests = grouped["Passed"].length;
    const passedPercent = totalTests ? ((passedTests / totalTests) * 100).toFixed(1) : 0;

    // Add improved cover page
    doc.addSection({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `${run.name}`, size: 48, bold: true })],
        }),
        new Paragraph({ text: "", spacing: { after: 300 } }),
        createParagraph(`Run ID: ${RUN_ID}`, 24),
        createParagraph(`Test Run Name: ${run.name}`, 24),
        createParagraph(`Associated JIRA: ${run.refs || 'N/A'}`, 24),
        createParagraph(`Generated On: ${new Date().toLocaleString()}`, 24),
        createParagraph(`Total Tests: ${totalTests}`, 24),
        createParagraph(`Passed: ${passedTests} (${passedPercent}%)`, 24),
      ],
    });

    // Add test sections
    const sectionChildren = [];
    for (const status of Object.keys(grouped)) {
      if (!grouped[status].length) continue;

      sectionChildren.push(createParagraph(`Status: ${status}`, 28, true));
      for (const test of grouped[status]) {
        const content = await processTest(test, status);
        sectionChildren.push(...content);
      }
    }

    // Add test content to document
    doc.addSection({ children: sectionChildren });
    await saveDocument();
  } catch (err) {
    console.error('Error fetching tests:', err.response?.data || err.message);
  }
}

// Create styled paragraph
// Create styled paragraph with custom font
function createParagraph(text, size = 24, bold = false, alignment = AlignmentType.LEFT) {
    return new Paragraph({
      alignment,
      spacing: { after: 200 },
      children: [
        new TextRun({
          text,
          size,
          bold,
          font: 'Calibri', // Set the font to Arial (or any other desired font)
        }),
      ],
    });
  }
  
  // Create a horizontal line
  function createLine() {
    return new Paragraph({
      border: { bottom: { color: 'auto', space: 1, value: BorderStyle.SINGLE, size: 6 } },
      spacing: { after: 300 },
    });
  }
  
  // Create table for test metadata (title first, then test ID and status)
  function createTestTable(testId, title, status) {
    const color = status_map[Object.keys(status_map).find(k => status_map[k].label === status)]?.color || "FFFFFF";
  
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { type: ShadingType.CLEAR, color: "auto", fill: color },
              children: [new Paragraph({ text: "Title", bold: true })],
            }),
            new TableCell({ children: [new Paragraph(title)] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              shading: { type: ShadingType.CLEAR, color: "auto", fill: color },
              children: [new Paragraph({ text: "Test ID", bold: true })],
            }),
            new TableCell({ children: [new Paragraph(testId.toString())] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              shading: { type: ShadingType.CLEAR, color: "auto", fill: color },
              children: [new Paragraph({ text: "Status", bold: true })],
            }),
            new TableCell({ children: [new Paragraph(status)] }),
          ],
        }),
      ],
    });
  }
  
  // Process and format a single test
async function processTest(test, status) {
    const children = [];
    const testId = test.id;
  
    // Add test title heading
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: test.title, size: 30, bold: true, font: 'Arial' })], // Use Arial font here
    }));
  
    // Add a visual line separator
    children.push(createLine());
  
    // Add the test metadata table
    children.push(createTestTable(testId, test.title, status));
  
    try {
      // Fetch and sanitize test results
      let results = (await axiosInstance.get(`/index.php?/api/v2/get_results/${testId}`)).data;
      if (results && results.results) results = results.results;
      results = Array.isArray(results) ? results.filter(r => r && (r.comment || r.attachments?.length)) : [];
  
      // Process each result
      for (const [i, result] of results.entries()) {
        children.push(createParagraph(`Result ${i + 1}`, 24, true, AlignmentType.LEFT));
  
        // Process and add comment and images (if any) in the correct order
        if (result.comment) {
          await processCommentImages(result.comment, testId, children);
        }
  
        // Render image attachments
        if (result.attachments?.length) {
          for (const att of result.attachments) {
            await processAttachment(att, testId, children);
          }
        }
  
        // Add spacing after each result
        children.push(new Paragraph({ spacing: { after: 300 }, children: [] }));
      }
    } catch (err) {
      console.error(`Failed to get results for Test ${testId}:`, err.response?.data || err.message);
    }
  
    // Add a line break after each test
    children.push(new Paragraph({ children: [new TextRun('')] }));
  
    return children;
  }
  
  
  
  
  // Extract image links from markdown-style image tags in comments
  async function processCommentImages(comment, testId, children) {
    const regex = /!\[\]\((index\.php\?\/attachments\/get\/[^\)]+)\)/g;
    let match;
    let lastPosition = 0;
  
    // This flag ensures we add text and images in the correct order
    let firstImageProcessed = false;
  
    while ((match = regex.exec(comment)) !== null) {
      const imageMarkdown = match[1];
      const relative = match[1];
      const id = relative.split('/').pop();
      const url = `${TESTRAIL_URL}/${relative}`;
      const filename = `comment_img_${id}.png`;
      const savePath = path.join(ATTACHMENTS_FOLDER, filename);
  
      try {
        const response = await axiosInstance.get(url, { responseType: 'arraybuffer' });
        const type = response.headers['content-type'];
        if (!type.startsWith('image/')) continue;
  
        // Save the image to disk
        fs.writeFileSync(savePath, response.data);
        const buffer = fs.readFileSync(savePath);
        const { width, height } = await getImageSize(savePath);
  
        // Add text before the image if any
        const textBeforeImage = comment.substring(lastPosition, match.index).trim();
        if (textBeforeImage) {
          children.push(new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: textBeforeImage, size: 22 })] }));
        }
  
        // Then, add the image
        children.push(new Paragraph({
          spacing: { after: 200 },
          children: [new ImageRun({
            data: buffer,
            transformation: {
              width: Math.min(width, MAX_IMAGE_WIDTH),
              height: Math.round((height / width) * Math.min(width, MAX_IMAGE_WIDTH)),
            },
          })],
        }));
  
        // Update the last position in the comment text
        lastPosition = regex.lastIndex;
  
        firstImageProcessed = true;
      } catch (err) {
        console.error(`Failed to download comment image ${filename} for Test ${testId}:`, err.message);
      }
    }
  
    // Add remaining text after the last image if any
    const remainingText = comment.substring(lastPosition).trim();
    if (remainingText) {
      children.push(new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: remainingText, size: 22 })] }));
    }
  }
  

// Helper function to get image size using sharp
async function getImageSize(filePath) {
  const { width, height } = await sharp(filePath).metadata();
  return { width, height };
}

// Save document as Word file
async function saveDocument() {
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT_DOC, buffer);
  console.log(`Document saved to ${OUTPUT_DOC}`);
}
