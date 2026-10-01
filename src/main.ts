/**
 * 대학 강의 시간표 생성 & 휴강 알림 시스템
 * (Vanilla TypeScript / DOM 기반 - No React)
 * - 학과, 교수님, 수업명, 강의실, 학점, 필수/선택(분반) 입력 및 관리
 * - 입력 과목을 한눈에 볼 수 있는 직관적인 목록 테이블
 * - 신속하고 안정적인 과목 삭제 기능 (iframe 친화적 인앱 모달 및 원클릭 삭제)
 * - 필수 고정 과목과 변동 가능 과목을 고려한 충돌 방지 자동 시간표 생성
 * - 실시간 휴강 공지 알림 (Web Audio 차임벨 + 브라우저 알림)
 * - 입력된 모든 과목 상세 정보 및 주간 매트릭스 엑셀(.xlsx) 파일 완벽 저장
 */

import * as XLSX from 'xlsx';
import './index.css';

// ==========================================
// 1. 데이터 타입 정의
// ==========================================

export type DayOfWeek = '월' | '화' | '수' | '목' | '금';

export interface TimeSlot {
  day: DayOfWeek;
  startPeriod: number; // 1 ~ 9
  duration: number;    // 1 ~ 4
  classroom: string;
}

export interface CourseItem {
  id: string;
  department: string;     // 학과
  courseName: string;     // 수업명
  professor: string;      // 교수님
  credits: number;        // 학점
  category: '전공필수' | '전공선택' | '교양필수' | '교양선택' | '일반교양'; // 이수구분
  isMandatory: boolean;   // 필수(고정) 과목 여부
  timeSlot: TimeSlot;     // 강의 시간 및 장소
  colorIndex: number;     // 테마 색상 인덱스
  cancellation?: {
    isCancelled: boolean;
    cancelDate: string;
    reason: string;
    makeUpDate?: string;
    notifiedAt: string;
  };
}

export interface NotificationLog {
  id: string;
  courseName: string;
  professor: string;
  date: string;
  reason: string;
  makeUpDate?: string;
  time: string;
}

// 교시별 표준 시간
export const PERIOD_TIMES = [
  { period: 1, time: '09:00 - 10:00' },
  { period: 2, time: '10:00 - 11:00' },
  { period: 3, time: '11:00 - 12:00' },
  { period: 4, time: '12:00 - 13:00' },
  { period: 5, time: '13:00 - 14:00' },
  { period: 6, time: '14:00 - 15:00' },
  { period: 7, time: '15:00 - 16:00' },
  { period: 8, time: '16:00 - 17:00' },
  { period: 9, time: '17:00 - 18:00' },
];

export const DAYS: DayOfWeek[] = ['월', '화', '수', '목', '금'];

// 깔끔하고 눈이 편안한 파스텔 테마 색상 팔레트
export const COLOR_PALETTES = [
  { bg: 'bg-blue-50', border: 'border-blue-300', text: 'text-blue-900', badge: 'bg-blue-100 text-blue-800' },
  { bg: 'bg-emerald-50', border: 'border-emerald-300', text: 'text-emerald-900', badge: 'bg-emerald-100 text-emerald-800' },
  { bg: 'bg-purple-50', border: 'border-purple-300', text: 'text-purple-900', badge: 'bg-purple-100 text-purple-800' },
  { bg: 'bg-amber-50', border: 'border-amber-300', text: 'text-amber-900', badge: 'bg-amber-100 text-amber-800' },
  { bg: 'bg-indigo-50', border: 'border-indigo-300', text: 'text-indigo-900', badge: 'bg-indigo-100 text-indigo-800' },
  { bg: 'bg-rose-50', border: 'border-rose-300', text: 'text-rose-900', badge: 'bg-rose-100 text-rose-800' },
  { bg: 'bg-teal-50', border: 'border-teal-300', text: 'text-teal-900', badge: 'bg-teal-100 text-teal-800' },
  { bg: 'bg-orange-50', border: 'border-orange-300', text: 'text-orange-900', badge: 'bg-orange-100 text-orange-800' },
];

// ==========================================
// 2. 초기 프리셋 샘플 데이터
// ==========================================

const PRESET_DEPARTMENTS: Record<string, Omit<CourseItem, 'id'>[]> = {
  '컴퓨터공학과': [
    {
      department: '컴퓨터공학과',
      courseName: '프로그래밍기초 (C/C++)',
      professor: '김철수 교수',
      credits: 3,
      category: '전공필수',
      isMandatory: true,
      timeSlot: { day: '월', startPeriod: 2, duration: 2, classroom: 'IT관 301호' },
      colorIndex: 0,
    },
    {
      department: '컴퓨터공학과',
      courseName: '컴퓨터시스템개론',
      professor: '이민호 교수',
      credits: 3,
      category: '전공필수',
      isMandatory: true,
      timeSlot: { day: '화', startPeriod: 3, duration: 2, classroom: 'IT관 205호' },
      colorIndex: 1,
    },
    {
      department: '컴퓨터공학과',
      courseName: '이산구조론',
      professor: '박지영 교수',
      credits: 3,
      category: '전공선택',
      isMandatory: false,
      timeSlot: { day: '수', startPeriod: 2, duration: 2, classroom: '공학관 402호' },
      colorIndex: 2,
    },
    {
      department: '컴퓨터공학과',
      courseName: '대학실용영어 1',
      professor: 'John Smith 교수',
      credits: 2,
      category: '교양필수',
      isMandatory: false,
      timeSlot: { day: '목', startPeriod: 2, duration: 2, classroom: '인문관 104호' },
      colorIndex: 3,
    },
    {
      department: '컴퓨터공학과',
      courseName: '오픈소스 소프트웨어 실습',
      professor: '최성훈 교수',
      credits: 3,
      category: '전공선택',
      isMandatory: false,
      timeSlot: { day: '목', startPeriod: 5, duration: 2, classroom: 'SW실습실 102호' },
      colorIndex: 4,
    },
    {
      department: '컴퓨터공학과',
      courseName: '신입생 세미나 & 진로탐색',
      professor: '김철수 교수',
      credits: 1,
      category: '교양선택',
      isMandatory: true,
      timeSlot: { day: '화', startPeriod: 6, duration: 1, classroom: 'IT관 소강당' },
      colorIndex: 5,
      cancellation: {
        isCancelled: true,
        cancelDate: '2026-10-06',
        reason: '교수님 해외 학회(ICSE) 참석',
        makeUpDate: '2026-10-20 6교시',
        notifiedAt: '2026-10-01 09:30',
      },
    },
  ],
  '경영학과': [
    {
      department: '경영학과',
      courseName: '경영학원론',
      professor: '오재원 교수',
      credits: 3,
      category: '전공필수',
      isMandatory: true,
      timeSlot: { day: '월', startPeriod: 2, duration: 2, classroom: '경영관 101호' },
      colorIndex: 0,
    },
    {
      department: '경영학과',
      courseName: '회계원리',
      professor: '정수빈 교수',
      credits: 3,
      category: '전공필수',
      isMandatory: true,
      timeSlot: { day: '수', startPeriod: 3, duration: 2, classroom: '경영관 204호' },
      colorIndex: 1,
    },
    {
      department: '경영학과',
      courseName: '경제학입문',
      professor: '강동원 교수',
      credits: 3,
      category: '전공선택',
      isMandatory: false,
      timeSlot: { day: '화', startPeriod: 2, duration: 2, classroom: '사회과학관 301호' },
      colorIndex: 2,
    },
    {
      department: '경영학과',
      courseName: '비즈니스 커뮤니케이션',
      professor: 'Sarah Kim 교수',
      credits: 3,
      category: '교양필수',
      isMandatory: false,
      timeSlot: { day: '목', startPeriod: 4, duration: 2, classroom: '인문관 202호' },
      colorIndex: 3,
    },
    {
      department: '경영학과',
      courseName: '마케팅관리',
      professor: '윤서진 교수',
      credits: 3,
      category: '전공선택',
      isMandatory: false,
      timeSlot: { day: '목', startPeriod: 6, duration: 2, classroom: '경영관 305호' },
      colorIndex: 4,
    },
  ],
  '인공지능학과': [
    {
      department: '인공지능학과',
      courseName: '인공지능 개론',
      professor: '도현우 교수',
      credits: 3,
      category: '전공필수',
      isMandatory: true,
      timeSlot: { day: '월', startPeriod: 3, duration: 2, classroom: 'AI센터 401호' },
      colorIndex: 4,
    },
    {
      department: '인공지능학과',
      courseName: '파이썬 데이터프로그래밍',
      professor: '한승우 교수',
      credits: 3,
      category: '전공필수',
      isMandatory: true,
      timeSlot: { day: '수', startPeriod: 2, duration: 2, classroom: 'AI실습실 1호' },
      colorIndex: 1,
    },
    {
      department: '인공지능학과',
      courseName: 'AI를 위한 선형대수학',
      professor: '송지은 교수',
      credits: 3,
      category: '전공선택',
      isMandatory: false,
      timeSlot: { day: '화', startPeriod: 4, duration: 2, classroom: '자연계관 202호' },
      colorIndex: 2,
    },
    {
      department: '인공지능학과',
      courseName: '컴퓨팅사고와 문제해결',
      professor: '배준혁 교수',
      credits: 3,
      category: '교양필수',
      isMandatory: false,
      timeSlot: { day: '목', startPeriod: 3, duration: 2, classroom: 'IT관 101호' },
      colorIndex: 3,
    },
    {
      department: '인공지능학과',
      courseName: '디지털 윤리와 AI',
      professor: '신유나 교수',
      credits: 2,
      category: '교양선택',
      isMandatory: false,
      timeSlot: { day: '금', startPeriod: 2, duration: 2, classroom: '본관 대강당' },
      colorIndex: 0,
    },
  ],
};

// ==========================================
// 3. 상태 관리 (localStorage)
// ==========================================

const STORAGE_COURSES_KEY = 'college_timetable_courses_v4';
const STORAGE_LOGS_KEY = 'college_timetable_logs_v4';

let registeredCourses: CourseItem[] = [];
let notificationLogs: NotificationLog[] = [];
let activeFilter = 'ALL'; // 'ALL', 'MANDATORY', 'ELECTIVE', 'CANCELLED'

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_COURSES_KEY);
    if (raw) {
      registeredCourses = JSON.parse(raw);
    } else {
      loadPreset('컴퓨터공학과', false);
    }

    const logRaw = localStorage.getItem(STORAGE_LOGS_KEY);
    if (logRaw) {
      notificationLogs = JSON.parse(logRaw);
    } else {
      notificationLogs = [
        {
          id: 'log_1',
          courseName: '신입생 세미나 & 진로탐색',
          professor: '김철수 교수',
          date: '2026-10-06',
          reason: '교수님 해외 학회(ICSE) 참석',
          makeUpDate: '2026-10-20 6교시',
          time: '2026-10-01 09:30',
        },
      ];
      localStorage.setItem(STORAGE_LOGS_KEY, JSON.stringify(notificationLogs));
    }
  } catch (err) {
    console.error('Failed to load storage', err);
    loadPreset('컴퓨터공학과', false);
  }
}

function saveData() {
  try {
    localStorage.setItem(STORAGE_COURSES_KEY, JSON.stringify(registeredCourses));
    localStorage.setItem(STORAGE_LOGS_KEY, JSON.stringify(notificationLogs));
  } catch (err) {
    console.error('Failed to save storage', err);
  }
}

function loadPreset(deptName: string, notify = true) {
  const preset = PRESET_DEPARTMENTS[deptName] || PRESET_DEPARTMENTS['컴퓨터공학과'];
  registeredCourses = preset.map((item, idx) => ({
    ...item,
    id: 'c_' + Date.now() + '_' + idx + '_' + Math.random().toString(36).substring(2, 6),
  }));
  saveData();
  renderAll();
  if (notify) {
    showToast(`'${deptName}' 기본 추천 과목 목록이 로드되었습니다.`);
  }
}

// ==========================================
// 4. 충돌 검사 및 시간표 자동 생성
// ==========================================

function checkOverlap(slotA: TimeSlot, slotB: TimeSlot): boolean {
  if (slotA.day !== slotB.day) return false;
  const startA = slotA.startPeriod;
  const endA = slotA.startPeriod + slotA.duration - 1;
  const startB = slotB.startPeriod;
  const endB = slotB.startPeriod + slotB.duration - 1;
  return Math.max(startA, startB) <= Math.min(endA, endB);
}

function findConflictCourse(candidateSlot: TimeSlot, excludeId?: string): CourseItem | undefined {
  return registeredCourses.find(c => {
    if (excludeId && c.id === excludeId) return false;
    return checkOverlap(c.timeSlot, candidateSlot);
  });
}

function autoReorganizeSchedule(options: { freeDay?: DayOfWeek | '상관없음' }) {
  const mandatory = registeredCourses.filter(c => c.isMandatory);
  const electives = registeredCourses.filter(c => !c.isMandatory);

  for (let i = 0; i < mandatory.length; i++) {
    for (let j = i + 1; j < mandatory.length; j++) {
      if (checkOverlap(mandatory[i].timeSlot, mandatory[j].timeSlot)) {
        showToast(`[시간표 충돌] 필수 과목 '${mandatory[i].courseName}'과 '${mandatory[j].courseName}' 시간이 겹칩니다.`, true);
        return;
      }
    }
  }

  let daysPool: DayOfWeek[] = ['월', '화', '수', '목', '금'];
  if (options.freeDay && options.freeDay !== '상관없음') {
    daysPool = daysPool.filter(d => d !== options.freeDay);
  }

  const placedCourses: CourseItem[] = [...mandatory];

  for (const elective of electives) {
    const currentConflict = placedCourses.some(c => checkOverlap(c.timeSlot, elective.timeSlot));
    if (!currentConflict && (!options.freeDay || elective.timeSlot.day !== options.freeDay)) {
      placedCourses.push(elective);
      continue;
    }

    let placed = false;
    const shuffledDays = [...daysPool].sort(() => Math.random() - 0.5);
    const candidatePeriods = [2, 3, 4, 5, 6, 1, 7];

    for (const d of shuffledDays) {
      for (const p of candidatePeriods) {
        if (p + elective.timeSlot.duration - 1 > 9) continue;
        const testSlot: TimeSlot = {
          ...elective.timeSlot,
          day: d,
          startPeriod: p,
        };
        const hasOverlap = placedCourses.some(c => checkOverlap(c.timeSlot, testSlot));
        if (!hasOverlap) {
          elective.timeSlot = testSlot;
          placedCourses.push(elective);
          placed = true;
          break;
        }
      }
      if (placed) break;
    }

    if (!placed) {
      placedCourses.push(elective);
    }
  }

  registeredCourses = placedCourses;
  saveData();
  renderAll();
  showToast(`시간표 자동 배치가 완료되었습니다! (${options.freeDay && options.freeDay !== '상관없음' ? options.freeDay + '요일 공강 적용' : '최적 분산'})`);
}

// ==========================================
// 5. 엑셀 (.xlsx) 저장 및 다운로드 기능 (SheetJS)
// ==========================================

function exportToExcel() {
  try {
    const wb = XLSX.utils.book_new();

    // 시트 1: [입력과목상세목록]
    const inputCoursesData = registeredCourses.map((c, index) => {
      const startP = c.timeSlot.startPeriod;
      const endP = c.timeSlot.startPeriod + c.timeSlot.duration - 1;
      const startTime = PERIOD_TIMES[startP - 1]?.time.split(' - ')[0] || '';
      const endTime = PERIOD_TIMES[endP - 1]?.time.split(' - ')[1] || '';

      return {
        '순번': index + 1,
        '학과': c.department,
        '과목명': c.courseName,
        '담당교수': c.professor,
        '이수구분': c.category,
        '필수여부': c.isMandatory ? '필수(고정)' : '선택(유동)',
        '학점': c.credits,
        '강의요일': c.timeSlot.day + '요일',
        '강의교시': `${startP}교시 ~ ${endP}교시 (${c.timeSlot.duration}시간)`,
        '수업시간': `${startTime} ~ ${endTime}`,
        '강의실': c.timeSlot.classroom,
        '휴강상태': c.cancellation?.isCancelled ? '휴강' : '정상',
        '휴강일자': c.cancellation?.cancelDate || '-',
        '휴강사유': c.cancellation?.reason || '-',
        '보강일정': c.cancellation?.makeUpDate || '-',
        '최근공지시각': c.cancellation?.notifiedAt || '-',
      };
    });

    const wsCourses = XLSX.utils.json_to_sheet(inputCoursesData);
    wsCourses['!cols'] = [
      { wch: 6 },
      { wch: 16 },
      { wch: 24 },
      { wch: 14 },
      { wch: 12 },
      { wch: 12 },
      { wch: 8 },
      { wch: 10 },
      { wch: 18 },
      { wch: 16 },
      { wch: 16 },
      { wch: 10 },
      { wch: 14 },
      { wch: 22 },
      { wch: 20 },
      { wch: 18 },
    ];
    XLSX.utils.book_append_sheet(wb, wsCourses, '입력과목상세목록');

    // 시트 2: [주간시간표]
    const matrixRows: (string | number)[][] = [];
    matrixRows.push(['2026학년도 대학 주간 강의 시간표']);
    matrixRows.push(['저장일시', new Date().toLocaleString()]);
    matrixRows.push(['총 과목 수', `${registeredCourses.length}과목`, '총 학점', `${registeredCourses.reduce((sum, c) => sum + c.credits, 0)}학점`]);
    matrixRows.push(['']);
    matrixRows.push(['교시', '시간', '월요일', '화요일', '수요일', '목요일', '금요일']);

    for (const pt of PERIOD_TIMES) {
      const row: (string | number)[] = [`${pt.period}교시`, pt.time];

      for (const day of DAYS) {
        const found = registeredCourses.find(c => {
          return c.timeSlot.day === day &&
            pt.period >= c.timeSlot.startPeriod &&
            pt.period < c.timeSlot.startPeriod + c.timeSlot.duration;
        });

        if (found) {
          const cancelTag = found.cancellation?.isCancelled ? ' [⚠️휴강]' : '';
          row.push(`${found.courseName} (${found.professor}, ${found.timeSlot.classroom})${cancelTag}`);
        } else {
          row.push('-');
        }
      }
      matrixRows.push(row);
    }

    const wsMatrix = XLSX.utils.aoa_to_sheet(matrixRows);
    wsMatrix['!cols'] = [
      { wch: 10 },
      { wch: 16 },
      { wch: 28 },
      { wch: 28 },
      { wch: 28 },
      { wch: 28 },
      { wch: 28 },
    ];
    XLSX.utils.book_append_sheet(wb, wsMatrix, '주간시간표');

    // 시트 3: [휴강및보강내역]
    const cancelledList = registeredCourses.filter(c => c.cancellation?.isCancelled);
    const cancelData = cancelledList.length > 0
      ? cancelledList.map((c, i) => ({
          '순번': i + 1,
          '학과': c.department,
          '과목명': c.courseName,
          '교수명': c.professor,
          '정규시간': `${c.timeSlot.day}요일 ${c.timeSlot.startPeriod}~${c.timeSlot.startPeriod + c.timeSlot.duration - 1}교시`,
          '강의실': c.timeSlot.classroom,
          '휴강일자': c.cancellation?.cancelDate || '-',
          '휴강사유': c.cancellation?.reason || '-',
          '보강일정': c.cancellation?.makeUpDate || '-',
          '공지발송시각': c.cancellation?.notifiedAt || '-',
        }))
      : [{ '안내': '현재 등록된 휴강 내역이 없습니다.' }];

    const wsCancel = XLSX.utils.json_to_sheet(cancelData);
    XLSX.utils.book_append_sheet(wb, wsCancel, '휴강보강내역');

    const deptTag = registeredCourses.length > 0 ? registeredCourses[0].department.replace(/\s+/g, '') : '대학생';
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const fileName = `강의시간표_${deptTag}_${dateStr}.xlsx`;

    XLSX.writeFile(wb, fileName);
    showToast(`📗 '${fileName}' 엑셀 파일(입력과목목록 + 시간표)이 저장되었습니다!`);
  } catch (err) {
    console.error('Excel export error', err);
    showToast('엑셀 파일 생성 중 오류가 발생했습니다: ' + (err as Error).message, true);
  }
}

// ==========================================
// 6. 차임벨 사운드 & 토스트 & 모달 시스템
// ==========================================

function playChime() {
  try {
    const AudioCtx = window.AudioContext || (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    const notes = [523.25, 659.25, 783.99]; // C5, E5, G5
    notes.forEach((freq, idx) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(freq, ctx.currentTime + idx * 0.12);
      gain.gain.setValueAtTime(0, ctx.currentTime + idx * 0.12);
      gain.gain.linearRampToValueAtTime(0.18, ctx.currentTime + idx * 0.12 + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + idx * 0.12 + 0.35);

      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(ctx.currentTime + idx * 0.12);
      osc.stop(ctx.currentTime + idx * 0.12 + 0.4);
    });
  } catch (e) {
    console.log('Audio autoplay prevented', e);
  }
}

function showToast(message: string, isError = false) {
  const container = document.getElementById('toast-box');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `px-4 py-2.5 rounded-xl text-xs font-semibold shadow-xl border flex items-center gap-2 transform transition duration-200 translate-y-1 opacity-0 pointer-events-auto ${
    isError
      ? 'bg-rose-900 border-rose-700 text-rose-100'
      : 'bg-slate-900 border-slate-700 text-white'
  }`;
  toast.innerHTML = `<span>${isError ? '⚠️' : '🔔'}</span> <span>${message}</span>`;

  container.appendChild(toast);
  requestAnimationFrame(() => {
    toast.classList.remove('translate-y-1', 'opacity-0');
  });

  setTimeout(() => {
    toast.classList.add('translate-y-1', 'opacity-0');
    setTimeout(() => toast.remove(), 250);
  }, isError ? 4500 : 3200);
}

// ==========================================
// 7. 인앱 과목 삭제 & 휴강 모달 (iframe 샌드박스 완벽 대응)
// ==========================================

// 인앱 삭제 확인 모달
function promptDeleteCourse(courseId: string) {
  const item = registeredCourses.find(c => c.id === courseId);
  if (!item) return;

  const backdrop = document.getElementById('cancel-modal-backdrop');
  const content = document.getElementById('cancel-modal-content');
  if (!backdrop || !content) return;

  content.innerHTML = `
    <div class="p-6 text-center space-y-4">
      <div class="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center mx-auto text-xl font-bold">
        🗑️
      </div>
      <div>
        <h3 class="font-bold text-slate-900 text-base">과목 삭제 확인</h3>
        <p class="text-xs text-slate-600 mt-1.5 leading-relaxed">
          <strong>'${item.courseName}'</strong> (${item.professor}) 과목을<br>시간표와 수강 목록에서 완전히 삭제하시겠습니까?
        </p>
      </div>
      <div class="flex items-center justify-center gap-2 pt-2">
        <button id="btn-cancel-delete-modal" class="px-4 py-2 border rounded-xl hover:bg-slate-50 font-semibold text-slate-600 text-xs transition cursor-pointer">
          취소
        </button>
        <button id="btn-confirm-delete-modal" class="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-xl shadow-xs text-xs transition cursor-pointer">
          삭제하기
        </button>
      </div>
    </div>
  `;

  backdrop.classList.remove('hidden');

  const close = () => backdrop.classList.add('hidden');
  document.getElementById('btn-cancel-delete-modal')?.addEventListener('click', close);
  document.getElementById('btn-confirm-delete-modal')?.addEventListener('click', () => {
    registeredCourses = registeredCourses.filter(c => c.id !== courseId);
    saveData();
    close();
    renderAll();
    showToast(`🗑️ '${item.courseName}' 과목이 정상적으로 삭제되었습니다.`);
  });
}

// 전체 비우기 확인 모달
function promptClearAll() {
  const backdrop = document.getElementById('cancel-modal-backdrop');
  const content = document.getElementById('cancel-modal-content');
  if (!backdrop || !content) return;

  content.innerHTML = `
    <div class="p-6 text-center space-y-4">
      <div class="w-12 h-12 rounded-full bg-amber-100 text-amber-700 flex items-center justify-center mx-auto text-xl font-bold">
        ⚠️
      </div>
      <div>
        <h3 class="font-bold text-slate-900 text-base">시간표 전체 비우기</h3>
        <p class="text-xs text-slate-600 mt-1.5 leading-relaxed">
          등록된 모든 과목 (${registeredCourses.length}개)을 삭제하시겠습니까?<br>삭제 후에는 복구할 수 없습니다.
        </p>
      </div>
      <div class="flex items-center justify-center gap-2 pt-2">
        <button id="btn-cancel-clear-modal" class="px-4 py-2 border rounded-xl hover:bg-slate-50 font-semibold text-slate-600 text-xs transition cursor-pointer">
          취소
        </button>
        <button id="btn-confirm-clear-modal" class="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-xl shadow-xs text-xs transition cursor-pointer">
          모두 삭제하기
        </button>
      </div>
    </div>
  `;

  backdrop.classList.remove('hidden');
  const close = () => backdrop.classList.add('hidden');
  document.getElementById('btn-cancel-clear-modal')?.addEventListener('click', close);
  document.getElementById('btn-confirm-clear-modal')?.addEventListener('click', () => {
    registeredCourses = [];
    saveData();
    close();
    renderAll();
    showToast('모든 과목이 삭제되었습니다.');
  });
}

// 휴강 알림 모달
function openCancellationModal(courseId: string) {
  const course = registeredCourses.find(c => c.id === courseId);
  if (!course) return;

  const backdrop = document.getElementById('cancel-modal-backdrop');
  const content = document.getElementById('cancel-modal-content');
  if (!backdrop || !content) return;

  const isCancelled = course.cancellation?.isCancelled ?? false;
  const today = new Date().toISOString().split('T')[0];

  content.innerHTML = `
    <div class="p-5 space-y-4">
      <div class="flex items-center justify-between pb-3 border-b border-slate-200">
        <div>
          <h3 class="font-bold text-slate-900 text-base">휴강 알림 등록 및 발송</h3>
          <p class="text-xs text-slate-500 mt-0.5">${course.department} · ${course.courseName} (${course.professor})</p>
        </div>
        <button id="modal-close-x" class="text-slate-400 hover:text-slate-700 text-lg p-1 cursor-pointer">✕</button>
      </div>

      <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 flex items-center justify-between text-xs">
        <div>
          <span class="font-bold text-slate-800">${course.timeSlot.day}요일 ${course.timeSlot.startPeriod}~${course.timeSlot.startPeriod + course.timeSlot.duration - 1}교시</span>
          <span class="text-slate-500 ml-1">(${course.timeSlot.classroom})</span>
        </div>
        <span class="px-2 py-0.5 rounded font-bold ${isCancelled ? 'bg-rose-100 text-rose-700' : 'bg-emerald-100 text-emerald-700'}">
          ${isCancelled ? '현재 휴강중' : '정상 수업'}
        </span>
      </div>

      <form id="cancel-form" class="space-y-3 text-xs">
        <div>
          <label class="block font-semibold text-slate-700 mb-1">휴강 일자</label>
          <input type="date" id="cancel-input-date" value="${course.cancellation?.cancelDate || today}" required class="w-full px-3 py-2 border rounded-lg bg-white focus:ring-2 focus:ring-rose-500 focus:outline-none" />
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">휴강 사유</label>
          <input type="text" id="cancel-input-reason" value="${course.cancellation?.reason || '교수님 학회 참석'}" placeholder="예: 교수님 학회 참석, 학교 행사 등" required class="w-full px-3 py-2 border rounded-lg bg-white focus:ring-2 focus:ring-rose-500 focus:outline-none" />
          <div class="flex flex-wrap gap-1 mt-1.5">
            <button type="button" class="preset-chip text-[11px] px-2 py-0.5 bg-slate-100 hover:bg-slate-200 rounded text-slate-700 cursor-pointer" data-text="교수님 학회 참석">#학회 참석</button>
            <button type="button" class="preset-chip text-[11px] px-2 py-0.5 bg-slate-100 hover:bg-slate-200 rounded text-slate-700 cursor-pointer" data-text="교내 행사 및 축제">#교내 행사</button>
            <button type="button" class="preset-chip text-[11px] px-2 py-0.5 bg-slate-100 hover:bg-slate-200 rounded text-slate-700 cursor-pointer" data-text="중간고사 시험대비 자율학습">#자율 학습</button>
            <button type="button" class="preset-chip text-[11px] px-2 py-0.5 bg-slate-100 hover:bg-slate-200 rounded text-slate-700 cursor-pointer" data-text="개인 사정 및 병가">#개인 사정</button>
          </div>
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">보강 일정 (선택)</label>
          <input type="text" id="cancel-input-makeup" value="${course.cancellation?.makeUpDate || ''}" placeholder="예: 10월 20일(화) 6교시 보강" class="w-full px-3 py-2 border rounded-lg bg-white focus:ring-2 focus:ring-rose-500 focus:outline-none" />
        </div>

        <div class="pt-2 flex items-center justify-between border-t border-slate-100">
          <div class="flex items-center gap-2">
            <button type="button" id="btn-delete-course-from-modal" class="text-rose-600 hover:text-rose-700 hover:underline font-semibold cursor-pointer text-xs">
              🗑️ 과목 삭제
            </button>
            ${isCancelled ? `
              <span class="text-slate-300">|</span>
              <button type="button" id="btn-revert-cancel" class="text-slate-600 hover:underline font-semibold cursor-pointer text-xs">휴강 취소(정상화)</button>
            ` : ''}
          </div>
          
          <div class="flex gap-2">
            <button type="button" id="modal-close-cancel" class="px-3.5 py-1.5 border rounded-lg hover:bg-slate-50 font-semibold text-slate-600 cursor-pointer">닫기</button>
            <button type="submit" class="px-4 py-1.5 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-lg shadow-xs cursor-pointer">휴강 공지 발송</button>
          </div>
        </div>
      </form>
    </div>
  `;

  backdrop.classList.remove('hidden');

  const close = () => backdrop.classList.add('hidden');
  document.getElementById('modal-close-x')?.addEventListener('click', close);
  document.getElementById('modal-close-cancel')?.addEventListener('click', close);

  // 모달 안에서 바로 과목 삭제
  document.getElementById('btn-delete-course-from-modal')?.addEventListener('click', () => {
    close();
    promptDeleteCourse(courseId);
  });

  // 사유 칩 클릭
  content.querySelectorAll('.preset-chip').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const txt = (e.currentTarget as HTMLElement).dataset.text;
      const input = document.getElementById('cancel-input-reason') as HTMLInputElement;
      if (input && txt) input.value = txt;
    });
  });

  // 휴강 취소
  document.getElementById('btn-revert-cancel')?.addEventListener('click', () => {
    delete course.cancellation;
    saveData();
    close();
    renderAll();
    showToast(`'${course.courseName}' 과목이 정상 수업으로 복귀되었습니다.`);
  });

  // 폼 제출
  const form = document.getElementById('cancel-form') as HTMLFormElement;
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const date = (document.getElementById('cancel-input-date') as HTMLInputElement).value;
    const reason = (document.getElementById('cancel-input-reason') as HTMLInputElement).value.trim();
    const makeup = (document.getElementById('cancel-input-makeup') as HTMLInputElement).value.trim();
    const notifiedAt = new Date().toLocaleString();

    course.cancellation = {
      isCancelled: true,
      cancelDate: date,
      reason,
      makeUpDate: makeup || undefined,
      notifiedAt,
    };

    notificationLogs.unshift({
      id: 'log_' + Date.now(),
      courseName: course.courseName,
      professor: course.professor,
      date,
      reason,
      makeUpDate: makeup || undefined,
      time: notifiedAt,
    });
    if (notificationLogs.length > 20) notificationLogs.pop();

    saveData();
    close();
    renderAll();

    playChime();

    if ('Notification' in window && Notification.permission === 'granted') {
      new Notification(`🚨 [휴강 공지] ${course.courseName}`, {
        body: `${date} 휴강 안내\n사유: ${reason}${makeup ? ` (보강: ${makeup})` : ''}`,
      });
    } else if ('Notification' in window && Notification.permission !== 'denied') {
      Notification.requestPermission();
    }

    showToast(`📢 [휴강 공지 발송] '${course.courseName}' ${date} 휴강 알림이 등록되었습니다.`);
  });
}

// ==========================================
// 8. 화면 렌더링 함수들
// ==========================================

// 1) 주간 시간표 매트릭스 렌더링
function renderTimetable() {
  const container = document.getElementById('timetable-view');
  if (!container) return;

  let html = `
    <table class="w-full text-left border-collapse border border-slate-200">
      <thead>
        <tr class="bg-slate-100/90 text-slate-700 text-xs border-b border-slate-200">
          <th class="py-2 px-2 text-center w-24 border-r border-slate-200 font-bold">교시 / 시간</th>
          ${DAYS.map(day => `
            <th class="py-2 px-2 text-center border-r last:border-r-0 border-slate-200 font-bold">
              <span class="inline-block px-2 py-0.5 rounded text-xs ${day === '월' || day === '금' ? 'text-indigo-700 font-extrabold' : 'text-slate-800'}">
                ${day}요일
              </span>
            </th>
          `).join('')}
        </tr>
      </thead>
      <tbody>
  `;

  for (const pt of PERIOD_TIMES) {
    html += `
      <tr class="border-b border-slate-100 min-h-[58px]">
        <td class="py-1 px-2 text-center border-r border-slate-200 bg-slate-50/70">
          <div class="text-xs font-bold text-slate-800">${pt.period}교시</div>
          <div class="text-[10px] text-slate-400 font-mono">${pt.time.split(' ')[0]}</div>
        </td>
    `;

    for (const day of DAYS) {
      const match = registeredCourses.find(c => {
        return c.timeSlot.day === day &&
          pt.period >= c.timeSlot.startPeriod &&
          pt.period < c.timeSlot.startPeriod + c.timeSlot.duration;
      });

      if (match) {
        const isStart = match.timeSlot.startPeriod === pt.period;
        const isCancelled = match.cancellation?.isCancelled ?? false;
        const pal = COLOR_PALETTES[match.colorIndex % COLOR_PALETTES.length] || COLOR_PALETTES[0];

        if (isStart) {
          html += `
            <td class="p-1 border-r last:border-r-0 border-slate-200 align-top relative group" rowspan="${match.timeSlot.duration}">
              <div class="h-full rounded-xl p-2.5 border transition-all flex flex-col justify-between shadow-2xs relative ${
                isCancelled
                  ? 'bg-rose-50 border-rose-300 text-rose-950 ring-1 ring-rose-400'
                  : `${pal.bg} ${pal.border} ${pal.text}`
              }">
                <!-- 과목 삭제 ✕ 버튼 (시간표 카드 우상단) -->
                <button 
                  class="btn-timetable-delete absolute top-1.5 right-1.5 text-slate-400 hover:text-rose-600 p-0.5 rounded hover:bg-white/80 transition cursor-pointer"
                  data-id="${match.id}"
                  title="과목 삭제"
                >
                  ✕
                </button>

                <div>
                  <div class="flex items-center justify-between gap-1 pr-5">
                    <span class="text-[10px] font-bold px-1.5 py-0.2 rounded ${pal.badge}">
                      ${match.isMandatory ? '필수' : '선택'} · ${match.credits}학점
                    </span>
                    ${isCancelled ? `
                      <span class="bg-rose-600 text-white text-[10px] font-bold px-1.5 py-0.2 rounded animate-pulse">
                        휴강
                      </span>
                    ` : ''}
                  </div>
                  
                  <div class="font-bold text-xs sm:text-sm mt-1 leading-snug ${isCancelled ? 'line-through text-rose-800' : ''}">
                    ${match.courseName}
                  </div>
                  
                  <div class="text-[11px] opacity-80 mt-0.5">
                    ${match.professor} · ${match.timeSlot.classroom}
                  </div>
                </div>

                <div class="mt-2 pt-1 border-t border-black/5 flex items-center justify-between text-[11px]">
                  <span class="text-[10px] font-mono opacity-70">
                    ${match.timeSlot.startPeriod}~${match.timeSlot.startPeriod + match.timeSlot.duration - 1}교시
                  </span>

                  <button 
                    class="btn-trigger-cancel px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                      isCancelled 
                        ? 'bg-rose-600 text-white hover:bg-rose-700' 
                        : 'bg-white hover:bg-slate-100 text-slate-800 border border-slate-300 shadow-2xs'
                    }"
                    data-id="${match.id}"
                  >
                    ${isCancelled ? '휴강공지' : '휴강설정'}
                  </button>
                </div>

                ${isCancelled && match.cancellation?.reason ? `
                  <div class="mt-1 text-[10px] text-rose-700 truncate" title="${match.cancellation.reason}">
                    사유: ${match.cancellation.reason}
                  </div>
                ` : ''}
              </div>
            </td>
          `;
        }
      } else {
        html += `
          <td class="p-1 border-r last:border-r-0 border-slate-100 text-center text-slate-200 text-xs">
            <span class="opacity-20 hover:opacity-100 hover:text-slate-400 font-mono cursor-default">공강</span>
          </td>
        `;
      }
    }

    html += `</tr>`;
  }

  html += `</tbody></table>`;
  container.innerHTML = html;

  // 휴강 모달 트리거
  container.querySelectorAll('.btn-trigger-cancel').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = (e.currentTarget as HTMLElement).dataset.id;
      if (id) openCancellationModal(id);
    });
  });

  // 시간표 카드에서 바로 삭제
  container.querySelectorAll('.btn-timetable-delete').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const id = (e.currentTarget as HTMLElement).dataset.id;
      if (id) promptDeleteCourse(id);
    });
  });
}

// 2) [입력 과목 상세 목록] 렌더링
function renderCoursesTable() {
  const container = document.getElementById('courses-table-body');
  const countSpan = document.getElementById('table-course-count');
  if (!container || !countSpan) return;

  let filtered = registeredCourses;
  if (activeFilter === 'MANDATORY') filtered = registeredCourses.filter(c => c.isMandatory);
  if (activeFilter === 'ELECTIVE') filtered = registeredCourses.filter(c => !c.isMandatory);
  if (activeFilter === 'CANCELLED') filtered = registeredCourses.filter(c => c.cancellation?.isCancelled);

  countSpan.textContent = `총 ${filtered.length}과목`;

  if (filtered.length === 0) {
    container.innerHTML = `
      <tr>
        <td colspan="8" class="py-6 text-center text-slate-400 text-xs">
          등록된 과목이 없습니다. 우측의 [새 강의 추가] 폼 또는 상단 학과별 프리셋을 불러와보세요.
        </td>
      </tr>
    `;
    return;
  }

  container.innerHTML = filtered.map((course, idx) => {
    const isCancelled = course.cancellation?.isCancelled ?? false;
    const pal = COLOR_PALETTES[course.colorIndex % COLOR_PALETTES.length] || COLOR_PALETTES[0];

    return `
      <tr class="border-b border-slate-100 hover:bg-slate-50 text-xs">
        <td class="py-2.5 px-3 text-center text-slate-400 font-mono">${idx + 1}</td>
        <td class="py-2.5 px-3 font-medium text-slate-600">${course.department}</td>
        <td class="py-2.5 px-3">
          <div class="flex items-center gap-1.5 font-bold text-slate-900">
            <span class="w-2.5 h-2.5 rounded-full ${pal.bg} border ${pal.border}"></span>
            <span class="${isCancelled ? 'line-through text-slate-400' : ''}">${course.courseName}</span>
          </div>
        </td>
        <td class="py-2.5 px-3 text-slate-600">${course.professor}</td>
        <td class="py-2.5 px-3">
          <span class="px-2 py-0.5 rounded font-bold text-[11px] ${
            course.isMandatory ? 'bg-indigo-100 text-indigo-800' : 'bg-slate-100 text-slate-700'
          }">
            ${course.isMandatory ? '필수(고정)' : '선택(유동)'}
          </span>
          <span class="text-slate-400 ml-1">· ${course.credits}학점</span>
        </td>
        <td class="py-2.5 px-3 font-medium text-slate-700">
          ${course.timeSlot.day}요일 ${course.timeSlot.startPeriod}~${course.timeSlot.startPeriod + course.timeSlot.duration - 1}교시
          <span class="text-slate-400">(${course.timeSlot.classroom})</span>
        </td>
        <td class="py-2.5 px-3 text-center">
          <span class="px-2 py-0.5 rounded-full font-bold text-[11px] ${
            isCancelled ? 'bg-rose-100 text-rose-700 border border-rose-300' : 'bg-emerald-50 text-emerald-700'
          }">
            ${isCancelled ? '⚠️ 휴강' : '정상'}
          </span>
        </td>
        <td class="py-2.5 px-3 text-right space-x-1 whitespace-nowrap">
          <button 
            class="btn-table-cancel px-2 py-1 rounded text-xs font-semibold cursor-pointer ${
              isCancelled ? 'bg-rose-600 text-white' : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
            }"
            data-id="${course.id}"
          >
            ${isCancelled ? '휴강수정' : '휴강설정'}
          </button>
          <button 
            class="btn-table-delete text-rose-600 hover:text-white hover:bg-rose-600 px-2 py-1 text-xs font-semibold rounded border border-rose-200 transition cursor-pointer"
            data-id="${course.id}"
            title="과목 삭제"
          >
            삭제
          </button>
        </td>
      </tr>
    `;
  }).join('');

  // 휴강 모달
  container.querySelectorAll('.btn-table-cancel').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = (e.currentTarget as HTMLElement).dataset.id;
      if (id) openCancellationModal(id);
    });
  });

  // 목록에서 삭제 버튼 (인앱 모달을 통해 100% 작동)
  container.querySelectorAll('.btn-table-delete').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = (e.currentTarget as HTMLElement).dataset.id;
      if (id) promptDeleteCourse(id);
    });
  });
}

// 3) 상단 메트릭 요약 바
function renderSummary() {
  const totalCourses = registeredCourses.length;
  const totalCredits = registeredCourses.reduce((sum, c) => sum + c.credits, 0);
  const totalCancelled = registeredCourses.filter(c => c.cancellation?.isCancelled).length;

  const countBadge = document.getElementById('summary-courses');
  const creditBadge = document.getElementById('summary-credits');
  const cancelBadge = document.getElementById('summary-cancelled');

  if (countBadge) countBadge.textContent = `${totalCourses}과목`;
  if (creditBadge) creditBadge.textContent = `${totalCredits}학점`;
  if (cancelBadge) {
    cancelBadge.textContent = `${totalCancelled}건 휴강`;
    cancelBadge.className = totalCancelled > 0
      ? 'px-2.5 py-1 rounded-lg text-xs font-bold bg-rose-100 text-rose-700 border border-rose-200'
      : 'px-2.5 py-1 rounded-lg text-xs font-bold bg-emerald-100 text-emerald-700 border border-emerald-200';
  }
}

// 4) 최근 휴강 공지 알림 로그 목록
function renderLogs() {
  const container = document.getElementById('log-list');
  if (!container) return;

  if (notificationLogs.length === 0) {
    container.innerHTML = `<div class="text-xs text-slate-400 py-3 text-center">발송된 휴강 공지가 없습니다.</div>`;
    return;
  }

  container.innerHTML = notificationLogs.slice(0, 5).map(log => `
    <div class="p-2.5 rounded-xl border border-rose-100 bg-rose-50/40 text-xs space-y-1">
      <div class="flex items-center justify-between text-[11px] text-slate-500">
        <span class="font-bold text-rose-700">📢 휴강 공지</span>
        <span>${log.time}</span>
      </div>
      <div class="font-bold text-slate-800">${log.courseName} (${log.professor})</div>
      <div class="text-slate-600 text-[11px]">
        <strong>${log.date}</strong> 휴강 · 사유: ${log.reason}
        ${log.makeUpDate ? `<br><span class="text-indigo-700">보강일정: ${log.makeUpDate}</span>` : ''}
      </div>
    </div>
  `).join('');
}

function renderAll() {
  renderTimetable();
  renderCoursesTable();
  renderSummary();
  renderLogs();
}

// ==========================================
// 9. 초기 DOM 생성 및 이벤트 등록
// ==========================================

function initApp() {
  const root = document.getElementById('root');
  if (!root) return;

  root.innerHTML = `
    <!-- 상단 토스트 알림 컨테이너 -->
    <div id="toast-box" class="fixed top-4 right-4 z-50 flex flex-col gap-2 pointer-events-none"></div>

    <div class="max-w-7xl mx-auto px-4 sm:px-6 py-6 space-y-6">

      <!-- 1. 헤더: 타이틀, 통계 뱃지, 액션 버튼 -->
      <header class="bg-white rounded-2xl p-5 shadow-xs border border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div class="flex items-center gap-2">
            <span class="text-2xl">🎓</span>
            <h1 class="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
              대학 강의 시간표 & 휴강 알림
            </h1>
          </div>
          <p class="text-xs text-slate-500 mt-1">
            학과·교수님·수업 정보를 등록하여 주간 시간표를 편성하고, 엑셀(.xlsx) 파일로 저장 및 실시간 휴강 공지를 발송합니다.
          </p>
        </div>

        <div class="flex flex-wrap items-center gap-2.5">
          <!-- 통계 뱃지 -->
          <div class="flex items-center gap-1.5 bg-slate-50 p-1.5 rounded-xl border border-slate-200">
            <span id="summary-courses" class="px-2.5 py-1 rounded-lg text-xs font-bold bg-white text-slate-800 shadow-2xs border border-slate-100">0과목</span>
            <span id="summary-credits" class="px-2.5 py-1 rounded-lg text-xs font-bold bg-indigo-600 text-white">0학점</span>
            <span id="summary-cancelled" class="px-2.5 py-1 rounded-lg text-xs font-bold bg-emerald-100 text-emerald-700">0건 휴강</span>
          </div>

          <!-- 엑셀 저장 버튼 (핵심 기능) -->
          <button id="btn-export-excel" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs sm:text-sm rounded-xl shadow-xs transition active:scale-95 flex items-center gap-1.5 cursor-pointer">
            <span>📗</span>
            <span>엑셀(.xlsx) 저장</span>
          </button>
        </div>
      </header>

      <!-- 2. 시간표 제어 및 자동 생성 바 -->
      <section class="bg-slate-900 rounded-2xl p-4 text-white shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div class="flex items-center gap-2">
          <span class="text-amber-400 text-base">⚡</span>
          <span class="text-xs sm:text-sm font-bold">시간표 자동 최적화 & 추천 세트 불러오기</span>
        </div>

        <div class="flex flex-wrap items-center gap-2">
          <!-- 학과 샘플 프리셋 로드 -->
          <select id="select-preset-dept" class="px-3 py-1.5 rounded-lg bg-slate-800 text-slate-200 text-xs font-semibold border border-slate-700">
            <option value="컴퓨터공학과">💻 컴퓨터공학과 세트</option>
            <option value="경영학과">📊 경영학과 세트</option>
            <option value="인공지능학과">🤖 인공지능학과 세트</option>
          </select>
          <button id="btn-load-preset" class="px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-white text-xs font-semibold rounded-lg transition cursor-pointer">
            불러오기
          </button>

          <!-- 공강 선호 자동 배치 -->
          <select id="select-free-day-opt" class="px-3 py-1.5 rounded-lg bg-slate-800 text-slate-200 text-xs font-semibold border border-slate-700">
            <option value="상관없음">공강 희망: 없음</option>
            <option value="금">금요일 공강 선호</option>
            <option value="월">월요일 공강 선호</option>
          </select>
          <button id="btn-auto-schedule" class="px-3.5 py-1.5 bg-indigo-500 hover:bg-indigo-600 text-white text-xs font-bold rounded-lg shadow-xs transition cursor-pointer">
            🪄 자동 시간표 재배치
          </button>

          <!-- 전체 비우기 -->
          <button id="btn-clear-all" class="px-2.5 py-1.5 bg-white/10 hover:bg-rose-900/60 text-slate-300 hover:text-white text-xs rounded-lg transition cursor-pointer" title="모든 과목 초기화">
            전체 비우기
          </button>
        </div>
      </section>

      <!-- 3. 메인 그리드 레이아웃: 좌측 주간 시간표 / 우측 과목 추가 & 휴강 센터 -->
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        
        <!-- 좌측 컬럼 (7): 주간 시간표 매트릭스 -->
        <div class="lg:col-span-7 space-y-4">
          <div class="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-xs space-y-3">
            <div class="flex items-center justify-between">
              <div class="flex items-center gap-2">
                <span class="text-base">📅</span>
                <h2 class="font-bold text-slate-900 text-sm sm:text-base">주간 강의 시간표 (월~금)</h2>
              </div>
              <div class="flex items-center gap-3 text-xs text-slate-500">
                <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-blue-400 inline-block"></span>정상</span>
                <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-rose-500 inline-block"></span>휴강</span>
                <span class="text-slate-400">· 카드 우상단 ✕로 즉시 삭제</span>
              </div>
            </div>

            <!-- 주간 시간표 매트릭스 테이블 컨테이너 -->
            <div id="timetable-view" class="overflow-x-auto"></div>
          </div>
        </div>

        <!-- 우측 컬럼 (5): 새 강의 추가 폼 & 실시간 휴강 공지 피드 -->
        <div class="lg:col-span-5 space-y-4">
          
          <!-- 새 강의 직접 등록 폼 -->
          <div class="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-xs space-y-3">
            <div class="flex items-center justify-between pb-2 border-b border-slate-100">
              <div class="flex items-center gap-1.5">
                <span class="text-base">✍️</span>
                <h3 class="font-bold text-slate-900 text-sm">새 강의 추가하기</h3>
              </div>
              <span class="text-[11px] text-slate-400">시간 충돌 시 자동 감지</span>
            </div>

            <form id="form-add-course" class="space-y-3 text-xs">
              <!-- 학과 & 수업명 -->
              <div class="grid grid-cols-2 gap-2">
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">학과 <span class="text-rose-500">*</span></label>
                  <input type="text" id="in-dept" value="컴퓨터공학과" required class="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 focus:bg-white focus:ring-1 focus:ring-indigo-500 focus:outline-none" />
                </div>
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">수업명 <span class="text-rose-500">*</span></label>
                  <input type="text" id="in-name" placeholder="예: 알고리즘" required class="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 focus:bg-white focus:ring-1 focus:ring-indigo-500 focus:outline-none" />
                </div>
              </div>

              <!-- 교수님 & 강의실 -->
              <div class="grid grid-cols-2 gap-2">
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">담당 교수님 <span class="text-rose-500">*</span></label>
                  <input type="text" id="in-prof" placeholder="예: 김교수님" required class="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 focus:bg-white focus:ring-1 focus:ring-indigo-500 focus:outline-none" />
                </div>
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">강의실</label>
                  <input type="text" id="in-room" placeholder="예: 공학관 301호" value="제1공학관 201호" class="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 focus:bg-white focus:ring-1 focus:ring-indigo-500 focus:outline-none" />
                </div>
              </div>

              <!-- 요일 / 시작교시 / 연강시간 -->
              <div class="grid grid-cols-3 gap-2">
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">요일</label>
                  <select id="in-day" class="w-full px-2 py-1.5 border rounded-lg bg-slate-50 focus:bg-white">
                    <option value="월">월요일</option>
                    <option value="화">화요일</option>
                    <option value="수">수요일</option>
                    <option value="목">목요일</option>
                    <option value="금">금요일</option>
                  </select>
                </div>
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">시작 교시</label>
                  <select id="in-period" class="w-full px-2 py-1.5 border rounded-lg bg-slate-50 focus:bg-white">
                    ${PERIOD_TIMES.map(p => `<option value="${p.period}">${p.period}교시 (${p.time.split(' ')[0]})</option>`).join('')}
                  </select>
                </div>
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">연강 시간</label>
                  <select id="in-duration" class="w-full px-2 py-1.5 border rounded-lg bg-slate-50 focus:bg-white">
                    <option value="1">1시간</option>
                    <option value="2" selected>2시간</option>
                    <option value="3">3시간</option>
                  </select>
                </div>
              </div>

              <!-- 학점 & 이수구분 & 필수(고정) 여부 -->
              <div class="grid grid-cols-3 gap-2">
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">학점</label>
                  <select id="in-credits" class="w-full px-2 py-1.5 border rounded-lg bg-slate-50">
                    <option value="3" selected>3학점</option>
                    <option value="2">2학점</option>
                    <option value="1">1학점</option>
                    <option value="4">4학점</option>
                  </select>
                </div>
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">이수구분</label>
                  <select id="in-category" class="w-full px-2 py-1.5 border rounded-lg bg-slate-50">
                    <option value="전공필수" selected>전공필수</option>
                    <option value="전공선택">전공선택</option>
                    <option value="교양필수">교양필수</option>
                    <option value="교양선택">교양선택</option>
                    <option value="일반교양">일반교양</option>
                  </select>
                </div>
                <div>
                  <label class="block font-semibold text-slate-700 mb-1">과목 구분</label>
                  <select id="in-mandatory" class="w-full px-2 py-1.5 border rounded-lg bg-slate-50">
                    <option value="true">필수(고정)</option>
                    <option value="false">선택(유동)</option>
                  </select>
                </div>
              </div>

              <button type="submit" class="w-full py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow-xs transition active:scale-95 cursor-pointer">
                강의 등록하기
              </button>
            </form>
          </div>

          <!-- 최근 휴강 알림 로그 피드 -->
          <div class="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-xs space-y-3">
            <div class="flex items-center justify-between pb-2 border-b border-slate-100">
              <div class="flex items-center gap-1.5">
                <span class="text-base">📢</span>
                <h3 class="font-bold text-slate-900 text-sm">최근 발송된 휴강 공지</h3>
              </div>
              <button id="btn-sound-test" class="text-[11px] text-indigo-600 hover:underline font-semibold cursor-pointer">
                🔔 소리 테스트
              </button>
            </div>
            <div id="log-list" class="space-y-2 max-h-48 overflow-y-auto pr-1"></div>
          </div>

        </div>

      </div>

      <!-- 4. [입력 과목 상세 목록 테이블] (입력 과목을 한눈에 파악할 수 있는 핵심 섹션) -->
      <section class="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs space-y-3">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-100">
          <div>
            <div class="flex items-center gap-2">
              <span class="text-base">📋</span>
              <h2 class="font-bold text-slate-900 text-base">입력된 수강 과목 상세 목록</h2>
              <span id="table-course-count" class="text-xs bg-slate-100 text-slate-700 font-bold px-2 py-0.5 rounded">0과목</span>
            </div>
            <p class="text-xs text-slate-400 mt-0.5">
              입력된 모든 과목의 상세 내역입니다. [삭제] 버튼을 눌러 언제든 과목을 제거할 수 있습니다.
            </p>
          </div>

          <!-- 필터 탭 -->
          <div class="flex items-center gap-1 bg-slate-50 p-1 rounded-xl border border-slate-200 text-xs">
            <button class="filter-btn px-2.5 py-1 rounded-lg font-bold bg-white text-slate-800 shadow-2xs cursor-pointer" data-filter="ALL">전체</button>
            <button class="filter-btn px-2.5 py-1 rounded-lg text-slate-600 hover:text-slate-900 cursor-pointer" data-filter="MANDATORY">필수만</button>
            <button class="filter-btn px-2.5 py-1 rounded-lg text-slate-600 hover:text-slate-900 cursor-pointer" data-filter="ELECTIVE">선택만</button>
            <button class="filter-btn px-2.5 py-1 rounded-lg text-slate-600 hover:text-slate-900 cursor-pointer" data-filter="CANCELLED">휴강만</button>
          </div>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse min-w-[720px]">
            <thead>
              <tr class="bg-slate-50 text-slate-600 text-xs border-y border-slate-200">
                <th class="py-2.5 px-3 text-center w-12">No</th>
                <th class="py-2.5 px-3">학과</th>
                <th class="py-2.5 px-3">수업명</th>
                <th class="py-2.5 px-3">담당교수</th>
                <th class="py-2.5 px-3">구분 / 학점</th>
                <th class="py-2.5 px-3">강의시간 및 장소</th>
                <th class="py-2.5 px-3 text-center">휴강상태</th>
                <th class="py-2.5 px-3 text-right">관리</th>
              </tr>
            </thead>
            <tbody id="courses-table-body"></tbody>
          </table>
        </div>
      </section>

    </div>

    <!-- 인앱 확인 및 휴강 알림 모달 -->
    <div id="cancel-modal-backdrop" class="hidden fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
      <div id="cancel-modal-content" class="bg-white rounded-2xl shadow-2xl max-w-md w-full border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-150"></div>
    </div>
  `;

  // ==========================================
  // 이벤트 바인딩
  // ==========================================

  // 엑셀 내보내기 버튼
  document.getElementById('btn-export-excel')?.addEventListener('click', exportToExcel);

  // 사운드 테스트 버튼
  document.getElementById('btn-sound-test')?.addEventListener('click', () => {
    playChime();
    showToast('🔔 맑은 차임벨 알림음이 재생되었습니다.');
  });

  // 프리셋 로드
  document.getElementById('btn-load-preset')?.addEventListener('click', () => {
    const dept = (document.getElementById('select-preset-dept') as HTMLSelectElement).value;
    loadPreset(dept, true);
  });

  // 자동 시간표 재배치
  document.getElementById('btn-auto-schedule')?.addEventListener('click', () => {
    const freeDay = (document.getElementById('select-free-day-opt') as HTMLSelectElement).value as DayOfWeek | '상관없음';
    autoReorganizeSchedule({ freeDay });
  });

  // 전체 비우기 (인앱 모달로 안전하게 처리)
  document.getElementById('btn-clear-all')?.addEventListener('click', () => {
    if (registeredCourses.length === 0) {
      showToast('삭제할 등록 과목이 없습니다.');
      return;
    }
    promptClearAll();
  });

  // 필터 탭
  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const target = e.currentTarget as HTMLElement;
      activeFilter = target.dataset.filter || 'ALL';

      document.querySelectorAll('.filter-btn').forEach(b => {
        b.className = 'filter-btn px-2.5 py-1 rounded-lg text-slate-600 hover:text-slate-900 cursor-pointer';
      });
      target.className = 'filter-btn px-2.5 py-1 rounded-lg font-bold bg-white text-slate-800 shadow-2xs cursor-pointer';

      renderCoursesTable();
    });
  });

  // 새 강의 폼 추가
  const form = document.getElementById('form-add-course') as HTMLFormElement;
  form.addEventListener('submit', (e) => {
    e.preventDefault();

    const dept = (document.getElementById('in-dept') as HTMLInputElement).value.trim();
    const name = (document.getElementById('in-name') as HTMLInputElement).value.trim();
    const prof = (document.getElementById('in-prof') as HTMLInputElement).value.trim();
    const classroom = (document.getElementById('in-room') as HTMLInputElement).value.trim() || '미지정';
    const day = (document.getElementById('in-day') as HTMLSelectElement).value as DayOfWeek;
    const startPeriod = parseInt((document.getElementById('in-period') as HTMLSelectElement).value, 10);
    const duration = parseInt((document.getElementById('in-duration') as HTMLSelectElement).value, 10);
    const credits = parseInt((document.getElementById('in-credits') as HTMLSelectElement).value, 10);
    const category = (document.getElementById('in-category') as HTMLSelectElement).value as CourseItem['category'];
    const isMandatory = (document.getElementById('in-mandatory') as HTMLSelectElement).value === 'true';

    if (startPeriod + duration - 1 > 9) {
      showToast('선택한 시작 교시와 연강 시간이 정규 9교시를 초과합니다.', true);
      return;
    }

    const newSlot: TimeSlot = { day, startPeriod, duration, classroom };

    // 시간 충돌 검사
    const conflict = findConflictCourse(newSlot);
    if (conflict) {
      showToast(`⚠️ [시간표 충돌] ${conflict.timeSlot.day}요일 ${conflict.timeSlot.startPeriod}~${conflict.timeSlot.startPeriod + conflict.timeSlot.duration - 1}교시에 이미 '${conflict.courseName}'(${conflict.professor}) 수업이 있습니다!`, true);
      return;
    }

    const newCourse: CourseItem = {
      id: 'c_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      department: dept,
      courseName: name,
      professor: prof,
      credits,
      category,
      isMandatory,
      timeSlot: newSlot,
      colorIndex: registeredCourses.length % COLOR_PALETTES.length,
    };

    registeredCourses.push(newCourse);
    saveData();
    renderAll();

    showToast(`'${name}' 과목이 성공적으로 등록되었습니다.`);
    (document.getElementById('in-name') as HTMLInputElement).value = '';
    (document.getElementById('in-prof') as HTMLInputElement).value = '';
  });
}

function start() {
  loadData();
  initApp();
  renderAll();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', start);
} else {
  start();
}
