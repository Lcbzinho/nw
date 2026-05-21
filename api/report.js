const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const REPORT_FILE_NAME = 'Auditoria_NW_Executive_Silencode.docx';

module.exports = (req, res) => {
  try {
    const outputDir = path.join(os.tmpdir(), 'nw-executive-outputs');
    const reportPath = path.join(outputDir, REPORT_FILE_NAME);

    execFileSync(process.execPath, [path.join(process.cwd(), 'generate_report.js')], {
      env: {
        ...process.env,
        OUTPUT_DIR: outputDir,
      },
      stdio: 'pipe',
    });

    const reportBuffer = fs.readFileSync(reportPath);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${REPORT_FILE_NAME}"`);
    res.status(200).send(reportBuffer);
  } catch (error) {
    res.status(500).send(`Falha ao gerar o relatório: ${error.message}`);
  }
};