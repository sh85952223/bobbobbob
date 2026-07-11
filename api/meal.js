const SCHOOL_NAME = '율량중학교';
const REGION_CODE = 'M10'; // 충청북도교육청
const OPEN_NEIS_BASE = 'https://open.neis.go.kr/hub';
const LEGACY_HOST = 'stu.cbe.go.kr';
const REQUEST_TIMEOUT_MS = 9000;
const MIDDLE_SCHOOL_CODE = '3';

let schoolInfoCache = null;
let legacySchoolCache = null;

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    res