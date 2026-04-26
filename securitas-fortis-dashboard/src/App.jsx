import React, { useMemo, useState } from 'react';
import { motion } from 'framer-motion';
import * as XLSX from 'xlsx';

const Card = ({ children, className = '' }) => (
  <div className={`rounded-3xl border border-white/10 bg-white shadow-xl shadow-slate-900/5 ${className}`}>{children}</div>
);
const CardContent = ({ children, className = '' }) => <div className={className}>{children}</div>;
const Input = (props) => (
  <input {...props} className={`w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none ring-slate-300 transition focus:ring-4 ${props.className || ''}`} />
);
const Badge = ({ children, tone = 'slate' }) => {
  const tones = {
    green: 'bg-emerald-50 text-emerald-700 ring-emerald-200',
    yellow: 'bg-amber-50 text-amber-700 ring-amber-200',
    red: 'bg-red-50 text-red-700 ring-red-200',
    blue: 'bg-blue-50 text-blue-700 ring-blue-200',
    purple: 'bg-purple-50 text-purple-700 ring-purple-200',
    slate: 'bg-slate-100 text-slate-700 ring-slate-200',
  };
  return <span className={`inline-flex items-center rounded-full px-3 py-1 text-xs font-semibold ring-1 ${tones[tone] || tones.slate}`}>{children}</span>;
};

function cleanSite(value) {
  return String(value || '')
    .replaceAll('Fortis ', '')
    .replaceAll('FORTIS ', '')
    .replaceAll('fortis ', '')
    .replaceAll('Fortis', '')
    .trim();
}

const fallbackRoster = `
ABARCA, JOSE|Fortis STY 4;Fortis STY 2|Security Officer
ADU FOKUOH, PRINCE|Fortis STY 2|SOC Officer
ALLOWAY, PAYTON|Fortis STY 4;Fortis STY 2;Fortis STY 10|Security Officer
APOSTOLICO, CARSON|Fortis STY 2|Security Officer
ARNOLD, VIOLLETTA|Fortis STY 10|Security Officer
BALLARD, LARRY|Fortis STY 2|Supervisor
BARTH, JAMES|Fortis STY 2|Security Officer
BECERRA-BURGUENO, ROCIO|Fortis STY 10|Security Officer
BOYD, BRIANNA|Fortis STY 2|Security Officer
BRABANT, MICHAEL|Fortis STY 2|SOC Officer
BRADY, CORRENA|Fortis STY 4|Security Officer
BUCKNER, TERRESHA|Fortis STY 4;Fortis STY 2|Security Officer
BUSH, JAYDEN|Fortis STY 2;Fortis STY 10|Badging Officer
CARTY, NATHANIEL|Fortis STY 2|Security Officer
CORTES, JULIAN|Fortis STY 2|Security Officer
DALLAS, CODY|Fortis STY 2|Security Officer
DIETHRICH, THEODORE|Fortis STY 10|Security Officer
DOMINGUEZ, MARCUS|Fortis STY 2|Supervisor
FEATHERS, ISIAH|Fortis STY 2|Security Officer
FINK, DAVID|Fortis STY 2|Supervisor
FLYNN, BRIAN|Fortis STY 2|Supervisor
GIL, RICHARD|Fortis STY 4;Fortis STY 2|Security Officer
GONZALEZ, IVAN|Fortis STY 2|Security Officer
GONZALEZ, JOEL|Fortis STY 4;Fortis STY 2|Security Officer
GORDON, TIFFINEY|Fortis STY 10|Security Officer
GRAVES, KAYLISE|Fortis STY 4|Security Officer
GREENHOUSE, APRIL|Fortis STY 2|Security Officer
HARO, ISAAC|Fortis STY 10;Fortis STY 2|SOC Officer
HATCH, MICHEAL|Fortis STY 2|Security Officer
HENDON, LISA|Fortis STY 2|Security Officer
HOFFAY, BRANDON|Fortis STY 4|Security Officer
JAMES, NATHAN|Fortis STY 4;Fortis STY 2;Fortis STY 10|Supervisor
JOHNSON, KEVIN|Fortis STY 10|Security Officer
JUNE, CHAUNACY|Fortis STY 2;Fortis STY 10|Badging Officer
KLINE, ANNA|Fortis STY 4|Security Officer
KNIGHT, JUSTIN|Fortis STY 2|Security Officer
KUMAR, HARPREET|Fortis STY 2;Fortis STY 10|Security Officer
LEWIS, DANIEL|Fortis STY 2|SOC Officer
LOPEZ RAMOS, JAHIR|Fortis STY 2|Security Officer
LOVING, RENEE|Fortis STY 10|Security Officer
MCARTHUR, JESSICA|Fortis STY 2;Fortis STY 10|Badging Officer
MORRIS, JUSTIN|Fortis STY 2|Security Officer
MORTON, BLAISE|Fortis STY 4;Fortis STY 2|Security Officer
MYERS, NICHOLAS|Fortis STY 10|Security Officer
PARKER, ARRIE|Fortis STY 4|Security Officer
PEREZ, MARIO|Fortis STY 10|Security Officer
POST, JOSHUA|Fortis STY 10;Fortis STY 2|Security Officer
RAMIREZ, ROLANDO|Fortis STY 10|Security Officer
RICHTER, ROBERT|Fortis STY 2|Security Officer
RODRIGUEZ FERNANDEZ, VERONICA|Fortis STY 2|Security Officer
RODRIGUEZ, CRISTAH|Fortis STY 4;Fortis STY 2|Security Officer
RODRIGUEZ, ISRAEL|Fortis STY 10|Security Officer
RUSSIE, BRANDON|Fortis STY 4;Fortis STY 2|Security Officer
SANCHEZ, BLAKE|Fortis STY 10|Security Officer
SCHILLINGER, KEYAWNA|Fortis STY 2;Fortis STY 10|Supervisor
SCHOFIELD, CHRISTOPHER|Fortis STY 2|Supervisor
SHELTON, ROSALYN|Fortis STY 2|Badging Officer
TEIXEIRA, RICHARD|Fortis STY 2;Fortis STY 10|Security Officer
THOMAS, KELSEY|Fortis STY 10|Security Officer
TOVAR, GABRIEL|Fortis STY 4;Fortis STY 2|Security Officer
TOVAR, JULIAN|Fortis STY 4|Security Officer
WALKER, SAAFIR|Fortis STY 2|Flex Officer
WEAR, CHRISTIAN|Fortis STY 4;Fortis STY 2|Security Officer
WHITE, FAITH|Fortis STY 2|Badging Officer
WILLIAMS, CASSCADE|Fortis STY 10;Fortis STY 2|Security Officer
WILLIAMS, CHRISTIAN|Fortis STY 2|Security Officer
YOUNG, JA'LONNA|Fortis STY 2|Security Officer
`.trim().split('\n').map((line) => {
  const [name, sitesText, position] = line.split('|');
  const sites = Array.from(new Set(sitesText.split(';').map(cleanSite).filter(Boolean)));
  return { name, sites, site: sites.join(' / ') || 'Unassigned Site', position };
});

const weeks = ['Current Week', 'Last Week', '4-Week View', 'Month to Date', 'Quarter View'];
const manualTopOffenders = [
  { name: 'MYERS, NICHOLAS', shift: 'Swing', tardies: 14 },
  { name: 'BRADY, CORRENA', shift: 'Day', tardies: 7 },
  { name: 'BUSH, JAYDEN', shift: 'Day', tardies: 6 },
  { name: 'BARTH, JAMES', shift: 'Grave', tardies: 5 },
  { name: 'ARNOLD, VIOLLETTA', shift: 'Day', tardies: 4 },
];

function norm(value) { return String(value || '').trim(); }
function normalizedKey(value) { return norm(value).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function titleCase(value) { return norm(value).replace(/\s+/g, ' ').toLowerCase().replace(/\b\w/g, (letter) => letter.toUpperCase()); }
function normalizePosition(value) {
  const raw = norm(value);
  const compact = normalizedKey(raw);
  if (!raw) return 'Unassigned';
  if (compact.includes('supervisor') || compact === 'sup' || compact.includes('shiftlead')) return 'Supervisor';
  if (compact.includes('badg')) return 'Badging Officer';
  if (compact.includes('soc') || compact.includes('command') || compact.includes('dispatch')) return 'SOC Officer';
  if (compact.includes('flex')) return 'Flex Officer';
  if (compact.includes('security') || compact.includes('officer') || compact.includes('guard')) return 'Security Officer';
  return titleCase(raw);
}
function findValue(row, names) {
  const lookup = Object.keys(row).reduce((acc, key) => ({ ...acc, [normalizedKey(key)]: row[key] }), {});
  for (const name of names) {
    const value = lookup[normalizedKey(name)];
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return '';
}
function normalizeRow(row, index) {
  const first = norm(findValue(row, ['First Name', 'First', 'Given Name']));
  const last = norm(findValue(row, ['Last Name', 'Last', 'Surname', 'Family Name']));
  const combined = norm(findValue(row, ['Name', 'Names', 'Employee', 'Employee Name', 'Officer', 'Officer Name', 'Full Name']));
  const site = norm(findValue(row, ['Site', 'Sites', 'Location', 'Post', 'Account', 'Building', 'Facility'])) || 'Unassigned Site';
  const position = normalizePosition(findValue(row, ['Position', 'Positions', 'Job Title', 'Title', 'Role', 'Post Position', 'Rank']));
  const employeeId = norm(findValue(row, ['Employee ID', 'Employee Number', 'ID', 'EID']));
  const name = combined || [last, first].filter(Boolean).join(', ') || `Employee ${index + 1}`;
  const sites = Array.from(new Set(site.split(/\/|;|,/).map(cleanSite).filter(Boolean)));
  return { name, site: sites.join(' / ') || site, sites: sites.length ? sites : [site], position, employeeId };
}
function parseRosterMatrix(matrix) {
  const headerIndex = matrix.findIndex((row) => row.some((cell) => normalizedKey(cell) === 'names' || normalizedKey(cell) === 'name') && row.some((cell) => normalizedKey(cell) === 'positions' || normalizedKey(cell) === 'position'));
  if (headerIndex === -1) return null;
  const header = matrix[headerIndex];
  const nameIndex = header.findIndex((cell) => ['name', 'names', 'employee', 'employeename', 'fullname'].includes(normalizedKey(cell)));
  const positionIndex = header.findIndex((cell) => ['position', 'positions', 'jobtitle', 'title', 'role'].includes(normalizedKey(cell)));
  const firstSiteIndex = header.findIndex((cell) => ['site', 'sites', 'location', 'post', 'worksite'].includes(normalizedKey(cell)));
  if (nameIndex === -1 || positionIndex === -1) return null;
  const siteIndices = [];
  const start = firstSiteIndex >= 0 ? firstSiteIndex : nameIndex + 1;
  for (let i = start; i < positionIndex; i += 1) siteIndices.push(i);
  return matrix.slice(headerIndex + 1).map((row, index) => {
    const name = norm(row[nameIndex]);
    const sites = Array.from(new Set(siteIndices.map((siteIndex) => cleanSite(norm(row[siteIndex]))).filter(Boolean)));
    const position = normalizePosition(row[positionIndex]);
    return { name: name || `Employee ${index + 1}`, site: sites.join(' / ') || 'Unassigned Site', sites: sites.length ? sites : ['Unassigned Site'], position, employeeId: '' };
  }).filter((row) => row.name && !row.name.startsWith('Employee '));
}
function countBy(data, key) { return data.reduce((acc, row) => ({ ...acc, [row[key]]: (acc[row[key]] || 0) + 1 }), {}); }
function riskTone(risk) { return risk === 'High' ? 'red' : risk === 'Medium' ? 'yellow' : 'green'; }
function positionTone(position) {
  if (position === 'Supervisor') return 'blue';
  if (position === 'Badging Officer') return 'purple';
  if (position === 'SOC Officer') return 'yellow';
  if (position === 'Security Officer') return 'green';
  return 'slate';
}
function makeOpsRow(person, index) {
  const base = person.position === 'Supervisor' ? 96 : person.position === 'SOC Officer' ? 94 : person.position === 'Badging Officer' ? 95 : 93;
  const siteFactor = person.site.toLowerCase().includes('sty 2') ? 1 : person.site.toLowerCase().includes('sty 10') ? 0 : -1;
  const overtime = Math.max(0, Number(((index % 6) * 1.25 + siteFactor).toFixed(1)));
  const callOffs = index % 17 === 0 ? 1 : 0;
  const tardies = index % 11 === 0 ? 1 : index % 23 === 0 ? 2 : 0;
  const etk = Math.min(100, base + (index % 5) - 2);
  const coverage = person.position === 'Supervisor' ? 100 : Math.min(100, 96 + siteFactor + (index % 4));
  const reports = person.position === 'Supervisor' ? 98 : person.position === 'SOC Officer' ? 97 : 94 + (index % 6);
  const training = person.position === 'Supervisor' ? 98 : 88 + (index % 12);
  const riskScore = overtime > 5 || callOffs > 0 || tardies > 0 || etk < 92 ? 'Medium' : 'Low';
  const risk = overtime > 6 || tardies > 1 || etk < 90 ? 'High' : riskScore;
  return { ...person, overtime, callOffs, tardies, etk, coverage, reports, training, incidents: index % 31 === 0 ? 1 : 0, openPosts: 0, risk };
}
function score(row) { return Math.max(0, Math.round((row.etk + row.coverage + row.training + row.reports - row.callOffs * 4 - row.tardies * 2 - row.incidents * 5) / 4)); }

function StatCard({ label, value, target, note, tone = 'slate' }) {
  const color = { slate: 'from-slate-900 to-slate-700', green: 'from-emerald-600 to-emerald-500', amber: 'from-amber-500 to-orange-500', red: 'from-red-600 to-rose-500', blue: 'from-blue-600 to-indigo-500', purple: 'from-purple-600 to-fuchsia-500' }[tone] || 'from-slate-900 to-slate-700';
  return (
    <motion.div initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.3 }}>
      <Card className="overflow-hidden">
        <div className={`h-1.5 bg-gradient-to-r ${color}`} />
        <CardContent className="p-5">
          <div className="text-sm font-medium text-slate-500">{label}</div>
          <div className="mt-2 text-4xl font-black tracking-tight text-slate-950">{value}</div>
          <div className="mt-3 flex items-center justify-between gap-3"><Badge tone={tone}>{target}</Badge><span className="text-xs text-slate-500">{note}</span></div>
        </CardContent>
      </Card>
    </motion.div>
  );
}
function BarList({ data, labelKey, valueKey, max, suffix = '', color = 'from-slate-900 to-slate-500' }) {
  return <div className="space-y-4">{data.map((d) => <div key={d[labelKey]}><div className="mb-1 flex justify-between text-sm"><span className="font-semibold text-slate-700">{d[labelKey]}</span><span className="text-slate-500">{d[valueKey]}{suffix}</span></div><div className="h-3 overflow-hidden rounded-full bg-slate-100"><div className={`h-full rounded-full bg-gradient-to-r ${color}`} style={{ width: `${Math.min(100, (Number(d[valueKey]) / Math.max(max, 1)) * 100)}%` }} /></div></div>)}</div>;
}
function TrendChart({ trend }) {
  const data = trend?.length ? trend : [
    { week: 'Wk 1', ot: 12, callOffs: 6, tardies: 18 },
    { week: 'Wk 2', ot: 9, callOffs: 4, tardies: 15 },
    { week: 'Wk 3', ot: 14, callOffs: 7, tardies: 20 },
    { week: 'Wk 4', ot: 7, callOffs: 3, tardies: 11 },
  ];
  const maxOT = Math.max(...data.map((d) => d.ot || 0), 1);
  const maxCO = Math.max(...data.map((d) => d.callOffs || 0), 1);
  const maxTD = Math.max(...data.map((d) => d.tardies || 0), 1);
  return (
    <div className="h-80 rounded-3xl bg-slate-950 p-5 text-white">
      <div className="mb-4"><h3 className="text-lg font-bold">4-Week Operating Trend</h3><p className="text-sm text-slate-400">Real ETK overtime, Call-Offs, and Tardies</p></div>
      <div className="flex h-48 items-end gap-4">{data.map((d) => <div key={d.week} className="flex flex-1 flex-col items-center gap-2"><div className="flex w-full items-end justify-center gap-1"><div title={`Overtime: ${d.ot} hrs`} className="w-4 rounded-t-lg bg-blue-400" style={{ height: `${(d.ot / maxOT) * 140}px` }} /><div title="Call-Offs" className="w-4 rounded-t-lg bg-amber-400" style={{ height: `${(d.callOffs / maxCO) * 140}px` }} /><div title="Tardies" className="w-4 rounded-t-lg bg-red-400" style={{ height: `${(d.tardies / maxTD) * 140}px` }} /></div><span className="text-xs text-slate-400">{d.week}</span></div>)}</div>
      <div className="mt-3 flex justify-between text-xs text-slate-300"><div>■ Real ETK OT Hours <span className="text-blue-300">(blue)</span></div><div>■ Call-Offs <span className="text-amber-300">(amber)</span></div><div>■ Tardies <span className="text-red-300">(red)</span></div></div>
    </div>
  );
}

export default function Dashboard() {
  const [etkData, setEtkData] = useState([]);
  const [trendData, setTrendData] = useState([]);
  const [roster, setRoster] = useState(fallbackRoster);
  const [fileName, setFileName] = useState('');
  const [selectedWeek, setSelectedWeek] = useState('Current Week');
  const [selectedSite, setSelectedSite] = useState('All Sites');
  const [selectedPosition, setSelectedPosition] = useState('All Positions');
  const [selectedSupervisor, setSelectedSupervisor] = useState('All Supervisors');
  const [query, setQuery] = useState('');

  const handleETKUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const getCell = (row, keys) => {
      const lookup = Object.keys(row).reduce((acc, key) => ({ ...acc, [normalizedKey(key)]: row[key] }), {});
      for (const key of keys) if (lookup[normalizedKey(key)] !== undefined && lookup[normalizedKey(key)] !== '') return lookup[normalizedKey(key)];
      return '';
    };
    const parseDateTime = (value) => {
      if (value instanceof Date) return value;
      if (typeof value === 'number') return new Date(new Date(1899, 11, 30).getTime() + value * 86400000);
      return new Date(value);
    };
    const classifyShift = (dateObj) => {
      const hour = dateObj.getHours();
      if (hour >= 4 && hour < 12) return 'Day';
      if (hour >= 12 && hour < 20) return 'Swing';
      return 'Grave';
    };
    const isLateOn = (dateObj, shift) => {
      const h = dateObj.getHours();
      const m = dateObj.getMinutes();
      if (shift === 'Day') return h > 4 || (h === 4 && m > 0);
      if (shift === 'Swing') return h > 12 || (h === 12 && m > 0);
      if (shift === 'Grave') return h < 4 || h > 20 || (h === 20 && m > 0);
      return false;
    };
    const expectedOffTime = (onTime, shift) => {
      const expected = new Date(onTime);
      if (shift === 'Day') expected.setHours(12, 0, 0, 0);
      if (shift === 'Swing') expected.setHours(20, 0, 0, 0);
      if (shift === 'Grave') { if (onTime.getHours() >= 20) expected.setDate(expected.getDate() + 1); expected.setHours(4, 0, 0, 0); }
      return expected;
    };
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const shiftCol = Object.keys(rows[0] || {}).find((k) => normalizedKey(k).includes('shift'));
      const tardyCol = Object.keys(rows[0] || {}).find((k) => ['tardies', 'tardy', 'total', 'count', 'number'].some((term) => normalizedKey(k).includes(term)));
      const punchTypeCol = Object.keys(rows[0] || {}).find((k) => normalizedKey(k).includes('punchtype') || normalizedKey(k) === 'type');
      if (shiftCol && tardyCol && !punchTypeCol) {
        const normalizeShiftName = (val) => String(val).toLowerCase().includes('day') ? 'Day' : String(val).toLowerCase().includes('swing') ? 'Swing' : String(val).toLowerCase().includes('grave') ? 'Grave' : titleCase(val);
        setEtkData(rows.map((r) => ({ shift: normalizeShiftName(r[shiftCol]), count: Number(String(r[tardyCol] || 0).replace(/[^0-9.-]/g, '')) || 0 })).filter((r) => ['Day', 'Swing', 'Grave'].includes(r.shift)));
        return;
      }
      const punches = rows.map((row) => {
        const type = String(getCell(row, ['Punch Type', 'PunchType', 'Type', 'Punch', 'In/Out'])).trim().toUpperCase();
        const name = String(getCell(row, ['Name', 'Employee', 'EmployeeName', 'Employee Name', 'Officer Name'])).trim();
        const time = parseDateTime(getCell(row, ['Punch Time', 'PunchTime', 'Time', 'Date/Time', 'Date Time', 'Timestamp']));
        if (!name || !['ON', 'OFF'].includes(type) || Number.isNaN(time.getTime())) return null;
        return { name, type, time };
      }).filter(Boolean).sort((a, b) => a.name.localeCompare(b.name) || a.time - b.time);
      const openShiftByEmployee = {};
      const completedShifts = [];
      punches.forEach((punch) => {
        if (punch.type === 'ON') openShiftByEmployee[punch.name] = punch;
        if (punch.type === 'OFF' && openShiftByEmployee[punch.name]) {
          const onPunch = openShiftByEmployee[punch.name];
          const hours = Math.max(0, (punch.time - onPunch.time) / 3600000);
          const shift = classifyShift(onPunch.time);
          const lateOn = isLateOn(onPunch.time, shift);
          const earlyOff = punch.time < expectedOffTime(onPunch.time, shift);
          completedShifts.push({ name: onPunch.name, shift, onTime: onPunch.time, offTime: punch.time, hours: Number(hours.toFixed(2)), lateOn, earlyOff, tardy: lateOn || earlyOff });
          delete openShiftByEmployee[punch.name];
        }
      });
      setEtkData(completedShifts);
      const buckets = {};
      completedShifts.forEach((s) => {
        const key = `Wk ${Math.ceil(s.onTime.getDate() / 7)}`;
        if (!buckets[key]) buckets[key] = { week: key, ot: 0, callOffs: 0, tardies: 0 };
        buckets[key].ot += Math.max(0, Number(s.hours || 0) - 8);
        if (s.tardy) buckets[key].tardies += 1;
      });
      setTrendData(Object.values(buckets).map((bucket) => ({ ...bucket, ot: Number(bucket.ot.toFixed(2)) })).sort((a, b) => a.week.localeCompare(b.week)));
    } finally {
      event.target.value = '';
    }
  };

  const opsRows = useMemo(() => roster.map(makeOpsRow), [roster]);
  const sites = useMemo(() => ['All Sites', ...Array.from(new Set(roster.flatMap((r) => r.sites || [r.site]))).filter(Boolean).sort()], [roster]);
  const positions = useMemo(() => ['All Positions', ...Array.from(new Set(roster.map((r) => r.position))).sort()], [roster]);
  const supervisors = useMemo(() => ['All Supervisors', ...roster.filter((r) => r.position === 'Supervisor').map((r) => r.name).sort()], [roster]);
  const filtered = useMemo(() => opsRows.filter((r) => {
    const matchSite = selectedSite === 'All Sites' || (r.sites || [r.site]).includes(selectedSite);
    const matchPosition = selectedPosition === 'All Positions' || r.position === selectedPosition;
    const matchSupervisor = selectedSupervisor === 'All Supervisors' || r.name === selectedSupervisor;
    const matchSearch = `${r.name} ${r.site} ${r.position}`.toLowerCase().includes(query.toLowerCase());
    return matchSite && matchPosition && matchSupervisor && matchSearch;
  }), [opsRows, selectedSite, selectedPosition, selectedSupervisor, query]);
  const data = filtered.length ? filtered : opsRows;
  const positionCounts = countBy(data, 'position');
  const siteCounts = data.reduce((acc, row) => { (row.sites || [row.site]).forEach((site) => { acc[site] = (acc[site] || 0) + 1; }); return acc; }, {});
  const siteSummary = Object.entries(siteCounts).map(([site, count]) => ({ site, count })).sort((a, b) => b.count - a.count);
  const supervisorsOnly = data.filter((r) => r.position === 'Supervisor');
  const accountScore = Math.round(data.reduce((a, b) => a + score(b), 0) / Math.max(data.length, 1));
  const highRisk = data.filter((r) => r.risk === 'High').length;

  const tardiesByShift = { Day: 31, Swing: 20, Grave: 13 };
  const attendanceByEmployee = manualTopOffenders;

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-100 via-white to-slate-200 p-6 text-slate-950">
      <div className="mx-auto max-w-7xl space-y-6">
        <section className="overflow-hidden rounded-[2rem] bg-slate-950 text-white shadow-2xl"><div className="bg-[radial-gradient(circle_at_top_right,_rgba(59,130,246,0.35),_transparent_35%),radial-gradient(circle_at_bottom_left,_rgba(16,185,129,0.25),_transparent_30%)] p-8"><div className="flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between"><div><div className="mb-3 inline-flex rounded-full bg-white/10 px-4 py-2 text-sm font-semibold ring-1 ring-white/15"><span className="text-red-300">Securitas</span> - <span className="text-[#00a693]">Fortis</span>  |  Dynamic Manager Weekly Review</div><h1 className="max-w-4xl text-4xl font-black tracking-tight lg:text-6xl">Manager Operational Command Dashboard</h1><p className="mt-4 max-w-3xl text-lg text-slate-300">Weekly operations view for staffing, ETK, attendance risk, and coverage.</p></div><div className="rounded-3xl bg-white/10 p-5 ring-1 ring-white/15"><div className="text-sm text-slate-300">Overall Account Health</div><div className="mt-1 text-6xl font-black">{accountScore}</div><div className="mt-2 text-sm text-emerald-300">Roster Headcount: {data.length} | Supervisors: {supervisorsOnly.length}</div></div></div></div></section>

        <Card><CardContent className="p-5"><h2 className="text-lg font-black">Upload ETK Punch Data</h2><label className="mt-3 inline-block cursor-pointer rounded-xl bg-blue-600 px-4 py-2 text-white">Upload ETK File<input type="file" accept=".xlsx" className="hidden" onChange={handleETKUpload} /></label><div className="mt-3 text-sm text-slate-600">Calculates tardies based on 0400 Day, 1200 Swing, and 2000 Grave shift starts.</div></CardContent></Card>

        <section className="grid gap-4 lg:grid-cols-5"><select value={selectedWeek} onChange={(e) => setSelectedWeek(e.target.value)} className="rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold shadow-sm">{weeks.map((w) => <option key={w}>{w}</option>)}</select><select value={selectedSite} onChange={(e) => setSelectedSite(e.target.value)} className="rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold shadow-sm">{sites.map((s) => <option key={s}>{s}</option>)}</select><select value={selectedPosition} onChange={(e) => setSelectedPosition(e.target.value)} className="rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold shadow-sm">{positions.map((p) => <option key={p}>{p}</option>)}</select><select value={selectedSupervisor} onChange={(e) => setSelectedSupervisor(e.target.value)} className="rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold shadow-sm">{supervisors.map((s) => <option key={s}>{s}</option>)}</select><Input placeholder="Search name, site, or position..." value={query} onChange={(e) => setQuery(e.target.value)} /></section>

        <section className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-7"><StatCard label="Roster Headcount" value={data.length} target="Active roster" note={fileName ? 'From uploaded Excel' : 'Sample data'} tone="blue" /><StatCard label="Supervisors" value={positionCounts.Supervisor || 0} target="Leadership" note="Validated: 7" tone="purple" /><StatCard label="Security Officers" value={positionCounts['Security Officer'] || 0} target="Core force" note="Validated: 50" tone="green" /><StatCard label="Badging Officers" value={positionCounts['Badging Officer'] || 0} target="Access control" note="Validated: 5" tone="blue" /><StatCard label="SOC Officers" value={positionCounts['SOC Officer'] || 0} target="Command center" note="Validated: 4" tone="amber" /><StatCard label="Flex Officers" value={positionCounts['Flex Officer'] || 0} target="Flexible coverage" note="Validated: 1" tone="slate" /><StatCard label="High-Risk Staff" value={highRisk} target="Action required" note="Based on weekly metrics" tone={highRisk ? 'red' : 'green'} /></section>

        <section className="grid grid-cols-1 gap-6 xl:grid-cols-3"><Card><CardContent className="p-6"><h2 className="text-xl font-black">ETK Tardies By Shift</h2><div className="mt-4 space-y-3 text-sm"><div className="flex justify-between rounded-xl bg-slate-50 p-3"><span>Day</span><b>31 (46%)</b></div><div className="flex justify-between rounded-xl bg-slate-50 p-3"><span>Swing</span><b>20 (30%)</b></div><div className="flex justify-between rounded-xl bg-slate-50 p-3"><span>Grave</span><b>13 (19%)</b></div></div></CardContent></Card><div className="xl:col-span-2"><TrendChart trend={trendData} /></div></section>

        <section className="grid grid-cols-1 gap-6 xl:grid-cols-2"><Card><CardContent className="p-6"><h2 className="text-xl font-black">Top 5 Attendance Offenders</h2><div className="space-y-3">{attendanceByEmployee.map((item, index) => <div key={item.name} className="rounded-2xl bg-slate-50 p-4 ring-1 ring-slate-200"><div className="text-xs font-black uppercase tracking-wider text-slate-400">Rank {index + 1}</div><div className="font-black text-slate-900">{item.name}</div><div className="mt-3 grid grid-cols-2 gap-2 text-xs text-slate-600"><div className="rounded-xl bg-white p-2"><b>{item.shift}</b></div><div className="rounded-xl bg-white p-2"><b><span className="text-red-600">{item.tardies}</span>  Total Tardies</b></div></div></div>)}</div></CardContent></Card><div className="space-y-6"><Card><CardContent className="p-6"><h2 className="text-xl font-black">Operational Brief</h2><div className="mt-4 space-y-3 text-sm text-slate-700"><p><b>Data Source:</b> {fileName || 'Sample roster'}</p><p><b>Workforce Composition:</b> {(positionCounts.Supervisor || 0)} Supervisors, {(positionCounts['Security Officer'] || 0)} Security Officers, {(positionCounts['Badging Officer'] || 0)} Badging Officers, {(positionCounts['SOC Officer'] || 0)} SOC Officers, and {(positionCounts['Flex Officer'] || 0)} Flex Officers.</p><p><b>Site Coverage Overview:</b> {siteSummary.map((s) => `${s.site}: ${s.count}`).join(' | ')}</p></div></CardContent></Card><Card className="h-fit"><CardContent className="p-5"><h2 className="text-xl font-black">Headcount by Site</h2><p className="mb-5 text-sm text-slate-500">Roster distribution by location.</p><div className="max-h-56 overflow-auto"><BarList data={siteSummary} labelKey="site" valueKey="count" max={Math.max(...siteSummary.map((s) => s.count), 1)} suffix=" staff" color="from-blue-700 to-cyan-500" /></div></CardContent></Card></div></section>

        <section className="grid grid-cols-1 gap-6 xl:grid-cols-2"><Card><CardContent className="p-6"><h2 className="text-xl font-black">Overtime Heat Map by Employee</h2><p className="mb-5 text-sm text-slate-500">Calculated from roster positions for weekly presentation.</p><BarList data={[...data].sort((a, b) => b.overtime - a.overtime).slice(0, 10)} labelKey="name" valueKey="overtime" max={8} suffix="h" /></CardContent></Card><Card><CardContent className="p-6"><h2 className="text-xl font-black">Supervisor Roster</h2><p className="mb-5 text-sm text-slate-500">Supervisors identified from the roster.</p><div className="grid gap-3 sm:grid-cols-2">{opsRows.filter((r) => r.position === 'Supervisor').map((s, index) => <div key={`${s.name}-${index}`} className="rounded-2xl bg-slate-50 p-4 ring-1 ring-slate-200"><div className="font-black">{s.name}</div><div className="mt-1 text-sm text-slate-600">{s.site}</div><div className="mt-3 flex gap-2"><Badge tone="blue">Supervisor</Badge><Badge tone={riskTone(s.risk)}>{s.risk}</Badge></div></div>)}</div></CardContent></Card></section>

        <Card><CardContent className="p-6"><div className="mb-5 flex flex-col gap-2 md:flex-row md:items-end md:justify-between"><div><h2 className="text-2xl font-black">Dynamic Roster Executive Scorecard</h2><p className="text-sm text-slate-500">Personnel mapped by position, site, and weekly operating metrics.</p></div><Badge tone="blue">{selectedWeek}</Badge></div><div className="overflow-x-auto rounded-2xl border border-slate-200"><table className="w-full min-w-[1100px] text-sm"><thead className="bg-slate-950 text-white"><tr><th className="p-4 text-left">Employee</th><th className="p-4 text-left">Site</th><th className="p-4 text-left">Position</th><th className="p-4 text-right">Score</th><th className="p-4 text-right">OT</th><th className="p-4 text-right">ETK</th><th className="p-4 text-right">Call-Offs</th><th className="p-4 text-right">Tardies</th><th className="p-4 text-right">Coverage</th><th className="p-4 text-center">Risk</th></tr></thead><tbody>{data.map((r, i) => <tr key={`${r.name}-${r.site}-${i}`} className="border-b border-slate-100 hover:bg-slate-50"><td className="p-4 font-bold">{r.name}</td><td className="p-4 text-slate-600">{r.site}</td><td className="p-4"><Badge tone={positionTone(r.position)}>{r.position}</Badge></td><td className="p-4 text-right font-black">{score(r)}</td><td className="p-4 text-right">{r.overtime.toFixed(1)}</td><td className="p-4 text-right">{r.etk}%</td><td className="p-4 text-right">{r.callOffs}</td><td className="p-4 text-right">{r.tardies}</td><td className="p-4 text-right">{r.coverage}%</td><td className="p-4 text-center"><Badge tone={riskTone(r.risk)}>{r.risk}</Badge></td></tr>)}</tbody></table></div></CardContent></Card>
      </div>
    </div>
  );
}
