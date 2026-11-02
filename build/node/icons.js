const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const gi = require('react-icons/gi');
const md = require('react-icons/md');
const out = process.argv[2];
const list = {
  brain: gi.GiBrain, database: fa.FaDatabase, image: fa.FaRegImages, sliders: fa.FaSlidersH,
  layers: fa.FaLayerGroup, cogs: fa.FaCogs, chart: fa.FaChartBar, eye: fa.FaEye, check: fa.FaCheckCircle,
  doctor: fa.FaUserMd, user: fa.FaUser, users: fa.FaUsers, laptop: fa.FaLaptopCode, shield: fa.FaShieldAlt,
  bolt: fa.FaBolt, redo: fa.FaRedoAlt, python: fa.FaPython, github: fa.FaGithub, cloud: fa.FaCloud,
  chip: fa.FaMicrochip, file: fa.FaFileAlt, search: fa.FaSearch, target: fa.FaBullseye, flask: fa.FaFlask,
  lock: fa.FaLock, gauge: fa.FaTachometerAlt, diagram: fa.FaProjectDiagram, balance: fa.FaBalanceScale,
  warn: fa.FaExclamationTriangle, bulb: fa.FaLightbulb, comments: fa.FaComments, tools: fa.FaTools,
  upload: fa.FaUpload, clone: fa.FaClone, books: fa.FaBook, grad: fa.FaUserGraduate, hdd: fa.FaHdd,
  fire: fa.FaFire, filter: fa.FaFilter, crop: fa.FaCropAlt, sitemap: fa.FaSitemap, flag: fa.FaFlagCheckered,
  clipboard: fa.FaClipboardCheck, code: fa.FaCode, rocket: fa.FaRocket, microscope: fa.FaMicroscope,
  stream: fa.FaStream, random: fa.FaRandom, puzzle: fa.FaPuzzlePiece, clock: fa.FaClock, tasks: fa.FaTasks,
  map: fa.FaMapSigns, link: fa.FaLink, tag: fa.FaTag, mri: md.MdOutlineMonitorHeart,
};
const colors = { w: '#FFFFFF', g: '#007367', d: '#1F2D2B', b: '#A57C52' };
(async () => {
  for (const [name, Icon] of Object.entries(list)) {
    if (!Icon) { console.log('MISSING', name); continue; }
    for (const [ck, col] of Object.entries(colors)) {
      const svg = renderToStaticMarkup(React.createElement(Icon, { color: col, size: '512' }));
      await sharp(Buffer.from(svg)).resize(384, 384, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toFile(`${out}/${name}_${ck}.png`);
    }
  }
  console.log('done', Object.keys(list).length);
})();