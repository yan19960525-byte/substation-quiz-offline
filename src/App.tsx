"use client";

import { ChangeEvent, MouseEvent as ReactMouseEvent, PointerEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type QuestionType = "single" | "multiple" | "judgment";
type FontSize = "xsmall" | "small" | "medium";
type AnswerState = "correct" | "wrong";

type Option = {
  id: string;
  text: string;
};

type Question = {
  id: string;
  stem: string;
  type: QuestionType;
  options: Option[];
  correct: string[];
  explanation: string;
  chapter: string;
  difficulty: string;
};

type Settings = {
  shuffleQuestions: boolean;
  shuffleOptions: boolean;
  fontSize: FontSize;
  autoAdvanceCorrect: boolean;
  autoAdvanceDelay: 300 | 500;
};

type Stats = {
  answered: number;
  correct: number;
};

type QuestionBank = {
  id: string;
  name: string;
  questions: Question[];
  stats: Stats;
  wrongIds: string[];
  importedAt: string;
};

type PersistedState = {
  version: 2;
  banks: QuestionBank[];
  activeBankId: string;
  settings: Settings;
};

type LegacyState = {
  questions?: Question[];
  settings?: Partial<Settings>;
  stats?: Stats;
  wrongIds?: string[];
};

type SessionAnswer = {
  selected: string[];
  correct: boolean;
};

type Screen = "home" | "quiz" | "result";
type SwipeAnimation = "exit-next" | "exit-previous" | "enter-next" | "enter-previous";

const DB_NAME = "local-question-trainer-v2";
const LEGACY_DB_NAME = "local-question-trainer";
const STORE_NAME = "app-state";
const STATE_KEY = "primary";

const DEFAULT_QUESTIONS: Question[] = [
  {
    id: "sample-1",
    stem: "导入题库后，题目和学习记录主要保存在哪里？",
    type: "single",
    options: [
      { id: "A", text: "当前设备本地" },
      { id: "B", text: "远程云数据库" },
      { id: "C", text: "公共题库服务器" },
      { id: "D", text: "微信账号" },
    ],
    correct: ["A"],
    explanation: "题库和学习记录通过浏览器本地存储保存在当前设备中。",
    chapter: "功能说明",
    difficulty: "易",
  },
  {
    id: "sample-2",
    stem: "导入另一份 Excel 题库后，原来的题库会怎样？",
    type: "single",
    options: [
      { id: "A", text: "继续保留，可以随时切换" },
      { id: "B", text: "立即被替换" },
      { id: "C", text: "自动上传到网络" },
      { id: "D", text: "无法再次打开" },
    ],
    correct: ["A"],
    explanation: "第二版支持保存多份题库，导入新题库不会覆盖原题库。",
    chapter: "第二版功能",
    difficulty: "易",
  },
  {
    id: "sample-3",
    stem: "当前版本支持哪些题型？",
    type: "multiple",
    options: [
      { id: "A", text: "单选题" },
      { id: "B", text: "多选题" },
      { id: "C", text: "判断题" },
      { id: "D", text: "视频问答题" },
    ],
    correct: ["A", "B", "C"],
    explanation: "当前版本支持单选、多选和判断三种题型。",
    chapter: "功能测试题",
    difficulty: "易",
  },
  {
    id: "sample-4",
    stem: "开启选项乱序后，判断题的“正确/错误”也会乱序。",
    type: "judgment",
    options: [
      { id: "TRUE", text: "正确" },
      { id: "FALSE", text: "错误" },
    ],
    correct: ["FALSE"],
    explanation: "判断题不进行选项乱序；该功能只用于单选题和多选题。",
    chapter: "功能测试题",
    difficulty: "易",
  },
];

const DEFAULT_SETTINGS: Settings = {
  shuffleQuestions: false,
  shuffleOptions: false,
  fontSize: "medium",
  autoAdvanceCorrect: true,
  autoAdvanceDelay: 500,
};

const SAMPLE_BANK: QuestionBank = {
  id: "sample-bank",
  name: "功能示例题库",
  questions: DEFAULT_QUESTIONS,
  stats: { answered: 0, correct: 0 },
  wrongIds: [],
  importedAt: "2026-07-29T00:00:00.000Z",
};

const DEFAULT_STATE: PersistedState = {
  version: 2,
  banks: [SAMPLE_BANK],
  activeBankId: SAMPLE_BANK.id,
  settings: DEFAULT_SETTINGS,
};

function openDatabase(databaseName = DB_NAME): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(databaseName, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE_NAME)) {
        request.result.createObjectStore(STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readState(databaseName: string): Promise<unknown> {
  const database = await openDatabase(databaseName);
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, "readonly");
    const request = transaction.objectStore(STORE_NAME).get(STATE_KEY);
    request.onsuccess = () => resolve(request.result ?? null);
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => database.close();
  });
}

async function loadState(): Promise<unknown> {
  const current = await readState(DB_NAME);
  if (current) return current;
  return readState(LEGACY_DB_NAME);
}

async function saveState(state: PersistedState): Promise<void> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, "readwrite");
    transaction.objectStore(STORE_NAME).put(state, STATE_KEY);
    transaction.oncomplete = () => {
      database.close();
      resolve();
    };
    transaction.onerror = () => reject(transaction.error);
  });
}

function validSettings(value: Partial<Settings> | undefined): Settings {
  const fontSizes: FontSize[] = ["xsmall", "small", "medium"];
  return {
    ...DEFAULT_SETTINGS,
    ...value,
    fontSize: value?.fontSize && fontSizes.includes(value.fontSize) ? value.fontSize : DEFAULT_SETTINGS.fontSize,
    autoAdvanceDelay: value?.autoAdvanceDelay === 300 ? 300 : 500,
  };
}

function migrateState(saved: unknown): PersistedState {
  if (!saved || typeof saved !== "object") return DEFAULT_STATE;
  const candidate = saved as Partial<PersistedState> & LegacyState;

  if (Array.isArray(candidate.banks) && candidate.banks.length > 0) {
    const banks = candidate.banks.filter(
      (bank): bank is QuestionBank => Boolean(bank?.id && bank?.name && Array.isArray(bank.questions)),
    );
    if (banks.length > 0) {
      return {
        version: 2,
        banks: banks.map((bank) => ({
          ...bank,
          stats: bank.stats ?? { answered: 0, correct: 0 },
          wrongIds: Array.isArray(bank.wrongIds) ? bank.wrongIds : [],
          importedAt: bank.importedAt ?? new Date().toISOString(),
        })),
        activeBankId: banks.some((bank) => bank.id === candidate.activeBankId)
          ? candidate.activeBankId!
          : banks[0].id,
        settings: validSettings(candidate.settings),
      };
    }
  }

  if (Array.isArray(candidate.questions) && candidate.questions.length > 0) {
    const migratedBank: QuestionBank = {
      id: "migrated-v1-bank",
      name: "第一版题库",
      questions: candidate.questions,
      stats: candidate.stats ?? { answered: 0, correct: 0 },
      wrongIds: Array.isArray(candidate.wrongIds) ? candidate.wrongIds : [],
      importedAt: new Date().toISOString(),
    };
    return {
      version: 2,
      banks: [migratedBank],
      activeBankId: migratedBank.id,
      settings: validSettings(candidate.settings),
    };
  }

  return DEFAULT_STATE;
}

function shuffle<T>(items: T[]): T[] {
  const result = [...items];
  for (let index = result.length - 1; index > 0; index -= 1) {
    const randomIndex = Math.floor(Math.random() * (index + 1));
    [result[index], result[randomIndex]] = [result[randomIndex], result[index]];
  }
  return result;
}

function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .replace(/\s+/g, "")
    .replace(/[（(][^）)]*[）)]/g, "")
    .trim();
}

function normalizeType(value: unknown): QuestionType | null {
  const text = String(value ?? "").trim();
  if (text.includes("多选")) return "multiple";
  if (text.includes("判断")) return "judgment";
  if (text.includes("单选")) return "single";
  return null;
}

function judgmentAnswer(value: unknown): "TRUE" | "FALSE" | null {
  const text = String(value ?? "").trim().toLowerCase();
  if (["正确", "对", "√", "true", "1"].includes(text)) return "TRUE";
  if (["错误", "错", "×", "x", "false", "0"].includes(text)) return "FALSE";
  return null;
}

function hashText(text: string): string {
  let hash = 0;
  for (let index = 0; index < text.length; index += 1) {
    hash = (hash * 31 + text.charCodeAt(index)) | 0;
  }
  return Math.abs(hash).toString(36);
}

export function parseWorkbook(arrayBuffer: ArrayBuffer): {
  questions: Question[];
  skipped: number;
  ignoredPlaceholders: number;
} {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!firstSheet) throw new Error("工作簿中没有可读取的工作表。");

  const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  const headerIndex = rows.findIndex((row) => {
    const headers = row.map(normalizeHeader);
    return headers.some((item) => item.startsWith("题干")) && headers.some((item) => item.startsWith("题型"));
  });
  if (headerIndex < 0) throw new Error("没有找到包含“题干”和“题型”的表头。");

  const headers = rows[headerIndex].map(normalizeHeader);
  const columnOf = (prefix: string) => headers.findIndex((header) => header.startsWith(prefix));
  const stemColumn = columnOf("题干");
  const typeColumn = columnOf("题型");
  const answerColumn = columnOf("正确答案");
  const explanationColumn = columnOf("解析");
  const chapterColumn = columnOf("章节");
  const difficultyColumn = columnOf("难度");
  const optionColumns = headers
    .map((header, index) => ({ match: header.match(/^选项([A-H])/i), index }))
    .filter((item): item is { match: RegExpMatchArray; index: number } => Boolean(item.match))
    .map((item) => ({ letter: item.match[1].toUpperCase(), index: item.index }));

  if (answerColumn < 0 || optionColumns.length === 0) {
    throw new Error("表格缺少“正确答案”或“选项 A～H”列。");
  }

  let skipped = 0;
  let ignoredPlaceholders = 0;
  const questions: Question[] = [];

  rows.slice(headerIndex + 1).forEach((row, offset) => {
    const stem = String(row[stemColumn] ?? "").trim();
    const type = normalizeType(row[typeColumn]);
    if (!stem || !type) {
      if (row.some((cell) => String(cell ?? "").trim())) skipped += 1;
      return;
    }

    let options = optionColumns
      .map(({ letter, index }) => ({ id: letter, text: String(row[index] ?? "").trim() }))
      .filter((option) => option.text.length > 0);

    const placeholderOptions = options.filter((option) => option.id >= "E" && option.text === "21");
    if (placeholderOptions.length > 0) {
      ignoredPlaceholders += placeholderOptions.length;
      options = options.filter((option) => !(option.id >= "E" && option.text === "21"));
    }

    let correct: string[];
    if (type === "judgment") {
      const answer = judgmentAnswer(row[answerColumn]);
      if (!answer) {
        skipped += 1;
        return;
      }
      options = [
        { id: "TRUE", text: "正确" },
        { id: "FALSE", text: "错误" },
      ];
      correct = [answer];
    } else {
      correct = Array.from(new Set(String(row[answerColumn] ?? "").toUpperCase().match(/[A-H]/g) ?? []));
      const optionIds = new Set(options.map((option) => option.id));
      if (options.length < 2 || correct.length === 0 || correct.some((id) => !optionIds.has(id))) {
        skipped += 1;
        return;
      }
    }

    questions.push({
      id: `import-${offset + headerIndex + 2}-${hashText(stem)}`,
      stem,
      type,
      options,
      correct,
      explanation: String(row[explanationColumn] ?? "").trim(),
      chapter: String(row[chapterColumn] ?? "").trim(),
      difficulty: String(row[difficultyColumn] ?? "").trim(),
    });
  });

  if (questions.length === 0) throw new Error("没有找到可导入的单选、多选或判断题。");
  return { questions, skipped, ignoredPlaceholders };
}

function arraysMatch(first: string[], second: string[]): boolean {
  if (first.length !== second.length) return false;
  const expected = new Set(second);
  return first.every((item) => expected.has(item));
}

function typeLabel(type: QuestionType): string {
  if (type === "multiple") return "多选题";
  if (type === "judgment") return "判断题";
  return "单选题";
}

function bankNameFromFile(filename: string): string {
  return filename.replace(/\.(xlsx|xls)$/i, "").trim() || "未命名题库";
}

export default function Home() {
  const [state, setState] = useState<PersistedState>(DEFAULT_STATE);
  const [hydrated, setHydrated] = useState(false);
  const [screen, setScreen] = useState<Screen>("home");
  const [session, setSession] = useState<Question[]>([]);
  const [sessionBankId, setSessionBankId] = useState("");
  const [sessionAnswers, setSessionAnswers] = useState<Record<string, SessionAnswer>>({});
  const [questionIndex, setQuestionIndex] = useState(0);
  const [selected, setSelected] = useState<string[]>([]);
  const [homeTab, setHomeTab] = useState<"study" | "library" | "settings">("study");
  const [navigatorOpen, setNavigatorOpen] = useState(false);
  const [wrongListOpen, setWrongListOpen] = useState(false);
  const [swipeAnimation, setSwipeAnimation] = useState<SwipeAnimation | null>(null);
  const [notice, setNotice] = useState("已内置功能示例题库，可直接开始体验。");
  const excelInput = useRef<HTMLInputElement>(null);
  const backupInput = useRef<HTMLInputElement>(null);
  const touchStart = useRef<{ x: number; y: number } | null>(null);
  const swipeBlockedClick = useRef(false);
  const swipeTimer = useRef<number | null>(null);

  useEffect(() => {
    loadState()
      .then((saved) => setState(migrateState(saved)))
      .catch(() => setNotice("本地记录读取失败，本次仍可继续使用。"))
      .finally(() => setHydrated(true));

    if ("serviceWorker" in navigator) {
      const serviceWorkerUrl = new URL("sw.js", window.location.href);
      const scope = new URL("./", window.location.href).pathname;
      navigator.serviceWorker.register(serviceWorkerUrl, { scope }).catch(() => undefined);
    }
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    saveState(state).catch(() => setNotice("本地保存失败，请及时导出备份。"));
  }, [state, hydrated]);

  const activeBank = state.banks.find((bank) => bank.id === state.activeBankId) ?? state.banks[0];
  const wrongListBank = screen === "quiz"
    ? state.banks.find((bank) => bank.id === sessionBankId) ?? activeBank
    : activeBank;
  const wrongQuestions = wrongListBank
    ? wrongListBank.questions.filter((question) => wrongListBank.wrongIds.includes(question.id))
    : [];
  const current = session[questionIndex];
  const currentAnswer = current ? sessionAnswers[current.id] : undefined;
  const answered = Boolean(currentAnswer);
  const lastCorrect = currentAnswer?.correct ?? false;
  const sessionAnswerList = Object.values(sessionAnswers);
  const sessionAnswered = sessionAnswerList.length;
  const sessionCorrect = sessionAnswerList.filter((answer) => answer.correct).length;
  const accuracy = activeBank?.stats.answered
    ? Math.round((activeBank.stats.correct / activeBank.stats.answered) * 100)
    : 0;

  const answerLabels = useMemo(() => {
    if (!current) return "";
    return current.correct
      .map((id) => {
        const optionIndex = current.options.findIndex((option) => option.id === id);
        return current.type === "judgment"
          ? current.options[optionIndex]?.text
          : String.fromCharCode(65 + optionIndex);
      })
      .filter(Boolean)
      .join("、");
  }, [current]);

  useEffect(() => {
    if (screen !== "quiz" || !currentAnswer?.correct || !state.settings.autoAdvanceCorrect || swipeAnimation) return;
    const timer = window.setTimeout(() => moveAfterAnsweredQuestion(), state.settings.autoAdvanceDelay);
    return () => window.clearTimeout(timer);
  }, [screen, questionIndex, currentAnswer, state.settings.autoAdvanceCorrect, state.settings.autoAdvanceDelay, swipeAnimation]);

  useEffect(() => () => {
    if (swipeTimer.current !== null) window.clearTimeout(swipeTimer.current);
  }, []);

  function patchSettings(patch: Partial<Settings>) {
    setState((previous) => ({ ...previous, settings: { ...previous.settings, ...patch } }));
  }

  function cycleFontSize() {
    const order: FontSize[] = ["xsmall", "small", "medium"];
    const currentIndex = order.indexOf(state.settings.fontSize);
    patchSettings({ fontSize: order[(currentIndex + 1) % order.length] });
  }

  function selectBank(bankId: string) {
    setState((previous) => ({ ...previous, activeBankId: bankId }));
    const bank = state.banks.find((item) => item.id === bankId);
    if (bank) setNotice(`已选择“${bank.name}”。`);
  }

  function startQuiz(wrongOnly = false, firstQuestionId?: string) {
    if (!activeBank) {
      setNotice("请先导入题库。");
      return;
    }
    let questions = wrongOnly
      ? activeBank.questions.filter((question) => activeBank.wrongIds.includes(question.id))
      : [...activeBank.questions];
    if (questions.length === 0) {
      setNotice(wrongOnly ? "当前题库的错题本是空的。" : "当前题库没有可练习的题目。");
      return;
    }
    if (firstQuestionId) {
      const firstQuestion = questions.find((question) => question.id === firstQuestionId);
      const remaining = questions.filter((question) => question.id !== firstQuestionId);
      if (state.settings.shuffleQuestions) questions = firstQuestion ? [firstQuestion, ...shuffle(remaining)] : shuffle(questions);
      else if (firstQuestion) questions = [firstQuestion, ...remaining];
    } else if (state.settings.shuffleQuestions) questions = shuffle(questions);
    questions = questions.map((question) => ({
      ...question,
      options:
        state.settings.shuffleOptions && question.type !== "judgment"
          ? shuffle(question.options)
          : [...question.options],
    }));
    setSession(questions);
    setSessionBankId(activeBank.id);
    setSessionAnswers({});
    setQuestionIndex(0);
    setSelected([]);
    setNavigatorOpen(false);
    setWrongListOpen(false);
    setSwipeAnimation(null);
    setScreen("quiz");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function gradeAnswer(answer: string[]) {
    if (!current || currentAnswer || swipeBlockedClick.current) return;
    const isCorrect = arraysMatch(answer, current.correct);
    setSelected(answer);
    setSessionAnswers((previous) => ({
      ...previous,
      [current.id]: { selected: answer, correct: isCorrect },
    }));
    setState((previous) => ({
      ...previous,
      banks: previous.banks.map((bank) => {
        if (bank.id !== sessionBankId) return bank;
        const wrongIds = new Set(bank.wrongIds);
        if (isCorrect) wrongIds.delete(current.id);
        else wrongIds.add(current.id);
        return {
          ...bank,
          wrongIds: Array.from(wrongIds),
          stats: {
            answered: bank.stats.answered + 1,
            correct: bank.stats.correct + (isCorrect ? 1 : 0),
          },
        };
      }),
    }));
  }

  function chooseOption(optionId: string) {
    if (!current || currentAnswer || swipeBlockedClick.current) return;
    if (current.type === "multiple") {
      setSelected((previous) =>
        previous.includes(optionId)
          ? previous.filter((id) => id !== optionId)
          : [...previous, optionId],
      );
      return;
    }
    gradeAnswer([optionId]);
  }

  function goToQuestion(index: number) {
    if (index < 0 || index >= session.length) return;
    const target = session[index];
    setQuestionIndex(index);
    setSelected(sessionAnswers[target.id]?.selected ?? []);
    setNavigatorOpen(false);
    setWrongListOpen(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function animateSwipeTo(index: number, direction: "next" | "previous") {
    if (index < 0 || index >= session.length || swipeAnimation) return;
    setSwipeAnimation(direction === "next" ? "exit-next" : "exit-previous");
    if (swipeTimer.current !== null) window.clearTimeout(swipeTimer.current);
    swipeTimer.current = window.setTimeout(() => {
      goToQuestion(index);
      setSwipeAnimation(direction === "next" ? "enter-next" : "enter-previous");
      swipeTimer.current = window.setTimeout(() => {
        setSwipeAnimation(null);
        swipeTimer.current = null;
      }, 210);
    }, 150);
  }

  function moveAfterAnsweredQuestion() {
    if (questionIndex + 1 < session.length) {
      goToQuestion(questionIndex + 1);
      return;
    }
    const firstUnanswered = session.findIndex((question) => !sessionAnswers[question.id]);
    if (firstUnanswered >= 0) {
      goToQuestion(firstUnanswered);
      return;
    }
    setScreen("result");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function onPointerDown(event: PointerEvent<HTMLElement>) {
    if (navigatorOpen || wrongListOpen || swipeAnimation) return;
    touchStart.current = { x: event.clientX, y: event.clientY };
    swipeBlockedClick.current = false;
    event.currentTarget.setPointerCapture(event.pointerId);
  }

  function onPointerUp(event: PointerEvent<HTMLElement>) {
    if (!touchStart.current) return;
    const deltaX = event.clientX - touchStart.current.x;
    const deltaY = event.clientY - touchStart.current.y;
    touchStart.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
    if (Math.abs(deltaX) < 60 || Math.abs(deltaX) <= Math.abs(deltaY) * 1.25) return;
    swipeBlockedClick.current = true;
    window.setTimeout(() => { swipeBlockedClick.current = false; }, 400);
    if (deltaX < 0) animateSwipeTo(questionIndex + 1, "next");
    else animateSwipeTo(questionIndex - 1, "previous");
  }

  function onPointerCancel(event: PointerEvent<HTMLElement>) {
    touchStart.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
  }

  function blockClickAfterSwipe(event: ReactMouseEvent<HTMLElement>) {
    if (!swipeBlockedClick.current) return;
    event.preventDefault();
    event.stopPropagation();
  }

  async function importExcel(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const result = parseWorkbook(await file.arrayBuffer());
      const bank: QuestionBank = {
        id: `bank-${Date.now()}-${hashText(file.name)}`,
        name: bankNameFromFile(file.name),
        questions: result.questions,
        stats: { answered: 0, correct: 0 },
        wrongIds: [],
        importedAt: new Date().toISOString(),
      };
      setState((previous) => ({
        ...previous,
        banks: [...previous.banks, bank],
        activeBankId: bank.id,
      }));
      const details = [
        `已新增题库“${bank.name}”，共 ${result.questions.length} 道题`,
        result.skipped ? `跳过 ${result.skipped} 行` : "",
        result.ignoredPlaceholders ? `忽略 ${result.ignoredPlaceholders} 个模板占位值“21”` : "",
      ].filter(Boolean);
      setNotice(`${details.join("，")}。原有题库继续保留。`);
    } catch (error) {
      setNotice(error instanceof Error ? `导入失败：${error.message}` : "导入失败，请检查表格格式。");
    }
  }

  function deleteBank(bankId: string) {
    const bank = state.banks.find((item) => item.id === bankId);
    if (!bank) return;
    if (state.banks.length === 1) {
      setNotice("至少需要保留一份题库。");
      return;
    }
    if (!window.confirm(`确认删除题库“${bank.name}”吗？该题库的错题和统计也会一并删除。`)) return;
    setState((previous) => {
      const banks = previous.banks.filter((item) => item.id !== bankId);
      return {
        ...previous,
        banks,
        activeBankId: previous.activeBankId === bankId ? banks[0].id : previous.activeBankId,
      };
    });
    setNotice(`已删除题库“${bank.name}”。`);
  }

  function exportBackup() {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `背题软件第二版备份-${new Date().toISOString().slice(0, 10)}.json`;
    link.click();
    URL.revokeObjectURL(url);
    setNotice(`备份已导出，包含 ${state.banks.length} 份题库。`);
  }

  async function importBackup(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as LegacyState & Partial<PersistedState>;
      if (!parsed || typeof parsed !== "object" || (!Array.isArray(parsed.banks) && !Array.isArray(parsed.questions))) {
        throw new Error("文件内容不完整");
      }
      const restored = migrateState(parsed);
      if (!restored.banks.length) throw new Error("文件内容不完整");
      setState(restored);
      setNotice(`备份恢复完成，共 ${restored.banks.length} 份题库。`);
    } catch {
      setNotice("备份恢复失败：请选择由本软件导出的 JSON 备份。");
    }
  }

  const wrongListOverlay = wrongListOpen && wrongListBank && (
    <div className="navigator-backdrop" role="presentation" onClick={() => setWrongListOpen(false)}>
      <section className="question-navigator wrong-list-sheet" role="dialog" aria-modal="true" aria-label="错题清单" onClick={(event) => event.stopPropagation()}>
        <div className="navigator-header">
          <span>
            <strong>错题清单</strong>
            <small>{wrongListBank.name} · {wrongQuestions.length} 题</small>
          </span>
          <button onClick={() => setWrongListOpen(false)} aria-label="关闭错题清单">×</button>
        </div>
        {wrongQuestions.length > 0 ? (
          <>
            <div className="wrong-question-list">
              {wrongQuestions.map((question) => {
                const originalIndex = wrongListBank.questions.findIndex((item) => item.id === question.id) + 1;
                return (
                  <button key={question.id} onClick={() => startQuiz(true, question.id)} aria-label={`练习第 ${originalIndex} 题：${question.stem}`}>
                    <span>第 {originalIndex} 题 · {typeLabel(question.type)}</span>
                    <strong>{question.stem}</strong>
                  </button>
                );
              })}
            </div>
            <button className="solid-action wrong-all-action" onClick={() => startQuiz(true)}>练习全部错题</button>
          </>
        ) : (
          <p className="empty-wrong-list">当前题库还没有错题。</p>
        )}
      </section>
    </div>
  );

  if (!hydrated) {
    return (
      <main className="loading-screen" role="status">
        <div className="brand-mark">题</div>
        <p>正在读取本地题库…</p>
      </main>
    );
  }

  if (screen === "quiz" && current) {
    const progress = session.length ? (sessionAnswered / session.length) * 100 : 0;
    const fontLabel: Record<FontSize, string> = { xsmall: "小", small: "中", medium: "大" };
    const sessionWrong = sessionAnswered - sessionCorrect;
    const allAnswered = sessionAnswered === session.length;
    const bankWrongCount = state.banks.find((bank) => bank.id === sessionBankId)?.wrongIds.length ?? 0;
    return (
      <main
        className={`app-shell quiz-shell font-${state.settings.fontSize}`}
        onPointerDown={onPointerDown}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerCancel}
        onClickCapture={blockClickAfterSwipe}
      >
        <header className="quiz-header">
          <button className="back-button" onClick={() => setScreen("home")} aria-label="退出本次练习">‹</button>
          <strong className="quiz-title">答题</strong>
          <button className="font-quick-button" onClick={cycleFontSize} aria-label="切换字体大小">
            Aa·{fontLabel[state.settings.fontSize]}
          </button>
        </header>
        <div className="progress-track" aria-label={`已完成 ${sessionAnswered} / ${session.length}`}>
          <div className="progress-fill" style={{ width: `${progress}%` }} />
        </div>

        <section className={`question-card ${swipeAnimation ? `swipe-${swipeAnimation}` : ""}`}>
          <div className="question-meta">
            <span className="question-number">⚑ {questionIndex + 1} / {session.length}</span>
            <span className={`type-badge type-${current.type}`}>{typeLabel(current.type)}</span>
          </div>
          <h1>{current.stem}</h1>
          {current.chapter && <p className="chapter-line">{current.chapter}</p>}

          <div className="options-list" role={current.type === "multiple" ? "group" : "radiogroup"}>
            {current.options.map((option, optionIndex) => {
              const displaySelected = currentAnswer?.selected ?? selected;
              const isSelected = displaySelected.includes(option.id);
              const isCorrectOption = answered && current.correct.includes(option.id);
              const isWrongOption = answered && isSelected && !current.correct.includes(option.id);
              return (
                <button
                  key={option.id}
                  className={`option-button ${isSelected ? "selected" : ""} ${isCorrectOption ? "correct" : ""} ${isWrongOption ? "wrong" : ""}`}
                  onClick={() => chooseOption(option.id)}
                  disabled={answered}
                  aria-pressed={isSelected}
                >
                  <span className="option-key">
                    {current.type === "judgment" ? (option.id === "TRUE" ? "✓" : "×") : String.fromCharCode(65 + optionIndex)}
                  </span>
                  <span className="option-text">{option.text}</span>
                </button>
              );
            })}
          </div>

          {current.type === "multiple" && !answered && (
            <button className="primary-button submit-answer" disabled={selected.length === 0} onClick={() => gradeAnswer(selected)}>
              提交答案
            </button>
          )}

          {answered && (
            <div className={`answer-panel ${lastCorrect ? "answer-correct" : "answer-wrong"}`} aria-live="polite">
              <div className="answer-title">
                <span>{lastCorrect ? "✓" : "!"}</span>
                <strong>{lastCorrect ? "回答正确" : "回答错误"}</strong>
                {lastCorrect && state.settings.autoAdvanceCorrect && (
                  <small>{state.settings.autoAdvanceDelay / 1000} 秒后自动跳转</small>
                )}
              </div>
              <p>正确答案：{answerLabels}</p>
              <div className="explanation">
                <span>解析</span>
                <p>{current.explanation || "本题暂未提供解析。"}</p>
              </div>
              {(!lastCorrect || !state.settings.autoAdvanceCorrect) && (
                <button className="primary-button" onClick={moveAfterAnsweredQuestion}>
                  {questionIndex + 1 < session.length ? "下一题" : allAnswered ? "查看结果" : "返回未答题"}
                </button>
              )}
            </div>
          )}
          <p className="swipe-hint">左滑下一题 · 右滑上一题</p>
        </section>

        <div className="quiz-bottom-bar" aria-label="本次练习统计">
          <button className="wrong-list-button" onClick={() => setWrongListOpen(true)} disabled={bankWrongCount === 0} aria-label={`查看错题清单，共 ${bankWrongCount} 题`}>
            ☆ 错题 {bankWrongCount}
          </button>
          <span className="bottom-correct">● {sessionCorrect}</span>
          <span className="bottom-wrong">● {sessionWrong}</span>
          <button className="progress-button" onClick={() => setNavigatorOpen(true)} aria-label="打开题号导航">
            {sessionAnswered} / {session.length}⌃
          </button>
        </div>

        {navigatorOpen && (
          <div className="navigator-backdrop" role="presentation" onClick={() => setNavigatorOpen(false)}>
            <section className="question-navigator" role="dialog" aria-modal="true" aria-label="题号导航" onClick={(event) => event.stopPropagation()}>
              <div className="navigator-header">
                <strong>题目进度</strong>
                <button onClick={() => setNavigatorOpen(false)} aria-label="关闭">×</button>
              </div>
              <div className="navigator-legend">
                <span><i className="legend-correct" />答对</span>
                <span><i className="legend-wrong" />答错</span>
                <span><i className="legend-unanswered" />未答</span>
              </div>
              <div className="question-grid">
                {session.map((question, index) => {
                  const answer = sessionAnswers[question.id];
                  const status: AnswerState | "unanswered" = answer ? (answer.correct ? "correct" : "wrong") : "unanswered";
                  return (
                    <button
                      key={question.id}
                      className={`${status} ${index === questionIndex ? "current" : ""}`}
                      onClick={() => goToQuestion(index)}
                      aria-label={`第 ${index + 1} 题，${status === "correct" ? "答对" : status === "wrong" ? "答错" : "未答"}`}
                    >
                      {index + 1}
                    </button>
                  );
                })}
              </div>
            </section>
          </div>
        )}
        {wrongListOverlay}
      </main>
    );
  }

  if (screen === "result") {
    const percent = sessionAnswered ? Math.round((sessionCorrect / sessionAnswered) * 100) : 0;
    return (
      <main className="app-shell result-shell">
        <section className="result-card">
          <div className="result-icon">✓</div>
          <p className="eyebrow">本次练习完成</p>
          <h1>{percent}<small>%</small></h1>
          <p>答对 {sessionCorrect} 题，共作答 {sessionAnswered} 题</p>
          <div className="result-actions">
            <button className="primary-button" onClick={() => startQuiz(false)}>再练一次</button>
            <button className="secondary-button" onClick={() => setScreen("home")}>返回首页</button>
          </div>
        </section>
      </main>
    );
  }

  return (
    <main className={`app-shell home-shell font-${state.settings.fontSize}`}>
      <header className="minimal-header">
        <h1>变电站背题</h1>
        <span>第二版 · 仅本地</span>
      </header>

      <nav className="home-tabs" aria-label="主要功能">
        <button className={homeTab === "study" ? "active" : ""} onClick={() => setHomeTab("study")}>答题</button>
        <button className={homeTab === "library" ? "active" : ""} onClick={() => setHomeTab("library")}>题库</button>
        <button className={homeTab === "settings" ? "active" : ""} onClick={() => setHomeTab("settings")}>设置</button>
      </nav>

      {homeTab === "study" && activeBank && (
        <section className="plain-panel study-panel">
          <label className="bank-picker">
            <span>当前题库</span>
            <select value={activeBank.id} onChange={(event) => selectBank(event.target.value)}>
              {state.banks.map((bank) => <option key={bank.id} value={bank.id}>{bank.name}</option>)}
            </select>
          </label>
          <div className="study-summary">
            <span>题目总数</span>
            <strong>{activeBank.questions.length}</strong>
            <small>道题</small>
          </div>
          <button className="solid-action" onClick={() => startQuiz(false)}>开始答题</button>
          <button className="line-action" onClick={() => setWrongListOpen(true)} disabled={activeBank.wrongIds.length === 0}>
            查看错题（{activeBank.wrongIds.length}）
          </button>
          <div className="simple-stats">
            <span>已答 <b>{activeBank.stats.answered}</b></span>
            <span>正确率 <b>{accuracy}%</b></span>
            <span>题库 <b>{state.banks.length}</b></span>
          </div>
        </section>
      )}

      {homeTab === "library" && (
        <section className="plain-panel">
          <h2>我的题库（{state.banks.length}）</h2>
          <div className="bank-list">
            {state.banks.map((bank) => (
              <article key={bank.id} className={`bank-row ${bank.id === state.activeBankId ? "active" : ""}`}>
                <button className="bank-main" onClick={() => selectBank(bank.id)}>
                  <strong>{bank.name}</strong>
                  <small>{bank.questions.length} 题 · 错题 {bank.wrongIds.length}</small>
                </button>
                {bank.id === state.activeBankId ? <span className="active-bank-mark">当前</span> : <button className="use-bank" onClick={() => selectBank(bank.id)}>选择</button>}
                <button className="delete-bank" onClick={() => deleteBank(bank.id)} aria-label={`删除${bank.name}`}>删除</button>
              </article>
            ))}
          </div>
          <input ref={excelInput} className="visually-hidden" type="file" accept=".xlsx,.xls" onChange={importExcel} />
          <button className="solid-action import-action" onClick={() => excelInput.current?.click()}>＋ 导入另一份 Excel 题库</button>
          <p className="plain-description">新题库会单独保存，不会覆盖已有题库。文件只在当前设备解析，不会上传。</p>
          <div className="plain-notice" role="status">{notice}</div>
          <hr />
          <h2>全部题库备份</h2>
          <input ref={backupInput} className="visually-hidden" type="file" accept="application/json,.json" onChange={importBackup} />
          <div className="two-actions">
            <button className="line-action" onClick={exportBackup}>导出备份</button>
            <button className="line-action" onClick={() => backupInput.current?.click()}>恢复备份</button>
          </div>
        </section>
      )}

      {homeTab === "settings" && (
        <section className="plain-panel">
          <div className="setting-line">
            <span><strong>题目乱序</strong><small>每次练习重新排列</small></span>
            <input type="checkbox" checked={state.settings.shuffleQuestions} onChange={(event) => patchSettings({ shuffleQuestions: event.target.checked })} />
          </div>
          <div className="setting-line">
            <span><strong>选项乱序</strong><small>不影响判断题</small></span>
            <input type="checkbox" checked={state.settings.shuffleOptions} onChange={(event) => patchSettings({ shuffleOptions: event.target.checked })} />
          </div>
          <div className="setting-line auto-setting">
            <span><strong>答对自动下一题</strong><small>答错时仍停留在当前题</small></span>
            <input type="checkbox" checked={state.settings.autoAdvanceCorrect} onChange={(event) => patchSettings({ autoAdvanceCorrect: event.target.checked })} />
          </div>
          {state.settings.autoAdvanceCorrect && (
            <div className="font-setting delay-setting">
              <span><strong>跳题延时</strong><small>选择看清正确提示所需时间</small></span>
              <div className="font-options delay-options" role="group" aria-label="自动跳题延时">
                {([300, 500] as const).map((delay) => (
                  <button key={delay} className={state.settings.autoAdvanceDelay === delay ? "active" : ""} onClick={() => patchSettings({ autoAdvanceDelay: delay })}>
                    {delay / 1000}秒
                  </button>
                ))}
              </div>
            </div>
          )}
          <div className="font-setting">
            <span><strong>答题字体</strong><small>题目、选项和行距同步调整</small></span>
            <div className="font-options" role="group" aria-label="字体大小">
              {(["xsmall", "small", "medium"] as const).map((size) => (
                <button key={size} className={state.settings.fontSize === size ? "active" : ""} onClick={() => patchSettings({ fontSize: size })}>
                  {({ xsmall: "小", small: "中", medium: "大" } as const)[size]}
                </button>
              ))}
            </div>
          </div>
        </section>
      )}

      {wrongListOverlay}
      <p className="privacy-note">无账号 · 无云端 · 无数据上传</p>
    </main>
  );
}
