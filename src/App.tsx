"use client";

import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type QuestionType = "single" | "multiple" | "judgment";

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
  fontSize: "small" | "medium" | "large";
};

type Stats = {
  answered: number;
  correct: number;
};

type PersistedState = {
  questions: Question[];
  settings: Settings;
  stats: Stats;
  wrongIds: string[];
};

type Screen = "home" | "quiz" | "result";

const DB_NAME = "local-question-trainer";
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
    stem: "更换题库时，应当选择哪种文件？",
    type: "single",
    options: [
      { id: "A", text: "Excel 题库文件" },
      { id: "B", text: "照片文件" },
      { id: "C", text: "音频文件" },
      { id: "D", text: "视频文件" },
    ],
    correct: ["A"],
    explanation: "当前版本按照约定模板识别 .xlsx 或 .xls 题库。",
    chapter: "功能说明",
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
    explanation: "第一版支持单选、多选和判断三种题型。",
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

const DEFAULT_STATE: PersistedState = {
  questions: DEFAULT_QUESTIONS,
  settings: { shuffleQuestions: false, shuffleOptions: false, fontSize: "medium" },
  stats: { answered: 0, correct: 0 },
  wrongIds: [],
};

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE_NAME)) {
        request.result.createObjectStore(STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function loadState(): Promise<PersistedState | null> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, "readonly");
    const request = transaction.objectStore(STORE_NAME).get(STATE_KEY);
    request.onsuccess = () => resolve((request.result as PersistedState) ?? null);
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => database.close();
  });
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
  if (!firstSheet) throw new Error("工作簿中没有可读取的工作表。 ");

  const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  const headerIndex = rows.findIndex((row) => {
    const headers = row.map(normalizeHeader);
    return headers.some((item) => item.startsWith("题干")) && headers.some((item) => item.startsWith("题型"));
  });
  if (headerIndex < 0) throw new Error("没有找到包含“题干”和“题型”的表头。 ");

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
    throw new Error("表格缺少“正确答案”或“选项 A～H”列。 ");
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

    const placeholderOptions = options.filter(
      (option) => option.id >= "E" && option.text === "21",
    );
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
      correct = Array.from(
        new Set(String(row[answerColumn] ?? "").toUpperCase().match(/[A-H]/g) ?? []),
      );
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

  if (questions.length === 0) throw new Error("没有找到可导入的单选、多选或判断题。 ");
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

export default function Home() {
  const [state, setState] = useState<PersistedState>(DEFAULT_STATE);
  const [hydrated, setHydrated] = useState(false);
  const [screen, setScreen] = useState<Screen>("home");
  const [session, setSession] = useState<Question[]>([]);
  const [questionIndex, setQuestionIndex] = useState(0);
  const [selected, setSelected] = useState<string[]>([]);
  const [answered, setAnswered] = useState(false);
  const [lastCorrect, setLastCorrect] = useState(false);
  const [sessionCorrect, setSessionCorrect] = useState(0);
  const [homeTab, setHomeTab] = useState<"study" | "library" | "settings">("study");
  const [notice, setNotice] = useState("已内置 4 道功能示例题，可直接开始体验。 ");
  const excelInput = useRef<HTMLInputElement>(null);
  const backupInput = useRef<HTMLInputElement>(null);

  useEffect(() => {
    loadState()
      .then((saved) => {
        if (saved?.questions?.length) {
          setState({
            ...DEFAULT_STATE,
            ...saved,
            settings: { ...DEFAULT_STATE.settings, ...saved.settings },
          });
        }
      })
      .catch(() => setNotice("本地记录读取失败，本次仍可继续使用。 "))
      .finally(() => setHydrated(true));

    if ("serviceWorker" in navigator) {
      const serviceWorkerUrl = new URL("sw.js", window.location.href);
      const scope = new URL("./", window.location.href).pathname;
      navigator.serviceWorker.register(serviceWorkerUrl, { scope }).catch(() => undefined);
    }
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    saveState(state).catch(() => setNotice("本地保存失败，请及时导出备份。 "));
  }, [state, hydrated]);

  const current = session[questionIndex];
  const accuracy = state.stats.answered
    ? Math.round((state.stats.correct / state.stats.answered) * 100)
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

  function patchSettings(patch: Partial<Settings>) {
    setState((previous) => ({ ...previous, settings: { ...previous.settings, ...patch } }));
  }

  function cycleFontSize() {
    const order: Settings["fontSize"][] = ["small", "medium", "large"];
    const currentIndex = order.indexOf(state.settings.fontSize);
    patchSettings({ fontSize: order[(currentIndex + 1) % order.length] });
  }

  function startQuiz(wrongOnly = false) {
    let questions = wrongOnly
      ? state.questions.filter((question) => state.wrongIds.includes(question.id))
      : [...state.questions];
    if (questions.length === 0) {
      setNotice(wrongOnly ? "错题本目前是空的。 " : "请先导入题库。 ");
      return;
    }
    if (state.settings.shuffleQuestions) questions = shuffle(questions);
    questions = questions.map((question) => ({
      ...question,
      options:
        state.settings.shuffleOptions && question.type !== "judgment"
          ? shuffle(question.options)
          : [...question.options],
    }));
    setSession(questions);
    setQuestionIndex(0);
    setSelected([]);
    setAnswered(false);
    setLastCorrect(false);
    setSessionCorrect(0);
    setScreen("quiz");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function gradeAnswer(answer: string[]) {
    if (!current || answered) return;
    const isCorrect = arraysMatch(answer, current.correct);
    setSelected(answer);
    setAnswered(true);
    setLastCorrect(isCorrect);
    if (isCorrect) setSessionCorrect((value) => value + 1);
    setState((previous) => {
      const wrongIds = new Set(previous.wrongIds);
      if (isCorrect) wrongIds.delete(current.id);
      else wrongIds.add(current.id);
      return {
        ...previous,
        wrongIds: Array.from(wrongIds),
        stats: {
          answered: previous.stats.answered + 1,
          correct: previous.stats.correct + (isCorrect ? 1 : 0),
        },
      };
    });
  }

  function chooseOption(optionId: string) {
    if (!current || answered) return;
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

  function nextQuestion() {
    if (questionIndex + 1 >= session.length) {
      setScreen("result");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    setQuestionIndex((value) => value + 1);
    setSelected([]);
    setAnswered(false);
    setLastCorrect(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function importExcel(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const result = parseWorkbook(await file.arrayBuffer());
      setState((previous) => ({
        ...previous,
        questions: result.questions,
        stats: { answered: 0, correct: 0 },
        wrongIds: [],
      }));
      const details = [
        `已导入 ${result.questions.length} 道题`,
        result.skipped ? `跳过 ${result.skipped} 行` : "",
        result.ignoredPlaceholders ? `忽略 ${result.ignoredPlaceholders} 个模板占位值“21”` : "",
      ].filter(Boolean);
      setNotice(`${details.join("，")}。原题库已由新题库替换。 `);
    } catch (error) {
      setNotice(error instanceof Error ? `导入失败：${error.message}` : "导入失败，请检查表格格式。 ");
    }
  }

  function exportBackup() {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `背题软件备份-${new Date().toISOString().slice(0, 10)}.json`;
    link.click();
    URL.revokeObjectURL(url);
    setNotice("备份已导出，请在 iPhone“文件”中妥善保存。 ");
  }

  async function importBackup(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as PersistedState;
      if (!Array.isArray(parsed.questions) || !parsed.settings || !parsed.stats) {
        throw new Error("文件内容不完整");
      }
      setState({ ...parsed, wrongIds: Array.isArray(parsed.wrongIds) ? parsed.wrongIds : [] });
      setNotice(`备份恢复完成，共 ${parsed.questions.length} 道题。 `);
    } catch {
      setNotice("备份恢复失败：请选择由本软件导出的 JSON 备份。 ");
    }
  }

  if (!hydrated) {
    return (
      <main className="loading-screen" role="status">
        <div className="brand-mark">题</div>
        <p>正在读取本地题库…</p>
      </main>
    );
  }

  if (screen === "quiz" && current) {
    const progress = ((questionIndex + (answered ? 1 : 0)) / session.length) * 100;
    const sessionAnswered = questionIndex + (answered ? 1 : 0);
    const fontLabel = { small: "小", medium: "中", large: "大" }[state.settings.fontSize];
    return (
      <main className={`app-shell quiz-shell font-${state.settings.fontSize}`}>
        <header className="quiz-header">
          <button className="back-button" onClick={() => setScreen("home")} aria-label="退出本次练习">
            ‹
          </button>
          <strong className="quiz-title">答题</strong>
          <button className="font-quick-button" onClick={cycleFontSize} aria-label="切换字体大小">
            Aa·{fontLabel}
          </button>
        </header>
        <div className="progress-track" aria-label={`练习进度 ${Math.round(progress)}%`}>
          <div className="progress-fill" style={{ width: `${progress}%` }} />
        </div>

        <section className="question-card">
          <div className="question-meta">
            <span className="question-number">⚑ {questionIndex + 1} / {session.length}</span>
            <span className={`type-badge type-${current.type}`}>{typeLabel(current.type)}</span>
          </div>
          <h1>{current.stem}</h1>
          {current.chapter && <p className="chapter-line">{current.chapter}</p>}

          <div className="options-list" role={current.type === "multiple" ? "group" : "radiogroup"}>
            {current.options.map((option, optionIndex) => {
              const isSelected = selected.includes(option.id);
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
              </div>
              <p>正确答案：{answerLabels}</p>
              <div className="explanation">
                <span>解析</span>
                <p>{current.explanation || "本题暂未提供解析。"}</p>
              </div>
              <button className="primary-button" onClick={nextQuestion}>
                {questionIndex + 1 === session.length ? "查看结果" : "下一题"}
              </button>
            </div>
          )}
        </section>
        <div className="quiz-bottom-bar" aria-label="本次练习统计">
          <span>☆ 错题 {state.wrongIds.length}</span>
          <span className="bottom-correct">● {sessionCorrect}</span>
          <span className="bottom-wrong">● {sessionAnswered - sessionCorrect}</span>
          <span>{sessionAnswered} / {session.length}</span>
        </div>
      </main>
    );
  }

  if (screen === "result") {
    const percent = session.length ? Math.round((sessionCorrect / session.length) * 100) : 0;
    return (
      <main className="app-shell result-shell">
        <section className="result-card">
          <div className="result-icon">✓</div>
          <p className="eyebrow">本次练习完成</p>
          <h1>{percent}<small>%</small></h1>
          <p>答对 {sessionCorrect} 题，共 {session.length} 题</p>
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
        <span>仅本地</span>
      </header>

      <nav className="home-tabs" aria-label="主要功能">
        <button className={homeTab === "study" ? "active" : ""} onClick={() => setHomeTab("study")}>答题</button>
        <button className={homeTab === "library" ? "active" : ""} onClick={() => setHomeTab("library")}>题库</button>
        <button className={homeTab === "settings" ? "active" : ""} onClick={() => setHomeTab("settings")}>设置</button>
      </nav>

      {homeTab === "study" && (
        <section className="plain-panel study-panel">
          <div className="study-summary">
            <span>题库总数</span>
            <strong>{state.questions.length}</strong>
            <small>道题</small>
          </div>
          <button className="solid-action" onClick={() => startQuiz(false)}>开始答题</button>
          <button className="line-action" onClick={() => startQuiz(true)} disabled={state.wrongIds.length === 0}>
            错题重练（{state.wrongIds.length}）
          </button>
          <div className="simple-stats">
            <span>已答 <b>{state.stats.answered}</b></span>
            <span>正确率 <b>{accuracy}%</b></span>
            <span>错题 <b>{state.wrongIds.length}</b></span>
          </div>
        </section>
      )}

      {homeTab === "library" && (
        <section className="plain-panel">
          <h2>导入 Excel 题库</h2>
          <p className="plain-description">支持单选、多选、判断题。文件只在当前设备中解析，不会上传。</p>
          <input ref={excelInput} className="visually-hidden" type="file" accept=".xlsx,.xls" onChange={importExcel} />
          <button className="solid-action" onClick={() => excelInput.current?.click()}>选择 Excel 文件</button>
          <div className="plain-notice" role="status">{notice}</div>
          <hr />
          <h2>本地备份</h2>
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
          <div className="font-setting">
            <span><strong>答题字体</strong><small>也可在答题页右上角快速切换</small></span>
            <div className="font-options" role="group" aria-label="字体大小">
              {(["small", "medium", "large"] as const).map((size) => (
                <button key={size} className={state.settings.fontSize === size ? "active" : ""} onClick={() => patchSettings({ fontSize: size })}>
                  {{ small: "小", medium: "中", large: "大" }[size]}
                </button>
              ))}
            </div>
          </div>
        </section>
      )}

      <p className="privacy-note">无账号 · 无云端 · 无数据上传</p>
    </main>
  );
}
