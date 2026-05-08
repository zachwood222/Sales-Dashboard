const fs = require('fs');
const path = require('path');
const multer = require('multer');
const { randomUUID } = require('crypto');

const UPLOAD_DIR = path.join(__dirname, '..', '..', 'data', 'uploads');
const MAX_FILE_SIZE_BYTES = 5 * 1024 * 1024;
const PDF_MIME_TYPE = 'application/pdf';

fs.mkdirSync(UPLOAD_DIR, { recursive: true });

function sanitizeBaseName(fileName) {
  return path
    .basename(fileName, path.extname(fileName))
    .replace(/[^a-zA-Z0-9_-]/g, '-')
    .slice(0, 80) || 'ticket';
}

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const timestamp = new Date().toISOString().replace(/[.:]/g, '-');
    const safeName = sanitizeBaseName(file.originalname);
    cb(null, `${timestamp}-${safeName}.pdf`);
  },
});

function fileFilter(_req, file, cb) {
  if (file.mimetype !== PDF_MIME_TYPE) {
    const err = new Error('Unsupported media type. Only application/pdf is accepted.');
    err.code = 'UNSUPPORTED_FILE_TYPE';
    return cb(err);
  }
  cb(null, true);
}

const uploadMiddleware = multer({
  storage,
  fileFilter,
  limits: {
    fileSize: MAX_FILE_SIZE_BYTES,
    files: 1,
  },
}).single('ticket');

function ingestTicketUpload(file) {
  const parseErrors = [];

  if (!file) {
    parseErrors.push({ code: 'MISSING_FILE', message: 'No ticket file was uploaded.' });
  }

  const ingestStatus = parseErrors.length ? 'failed' : 'queued';

  return {
    uploadId: randomUUID(),
    filename: file ? file.filename : null,
    ingestStatus,
    parseErrors,
  };
}

module.exports = {
  MAX_FILE_SIZE_BYTES,
  PDF_MIME_TYPE,
  uploadMiddleware,
  ingestTicketUpload,
};
