import readline from "node:readline";
import { stdin, stdout } from "node:process";
import { styleText } from "node:util";

/**
 * Two prompts, written out rather than pulled in.
 *
 * A scaffolder is the one place where a dependency is pure cost: it runs once,
 * through `npm create`, on a machine that has to download it first. Everything
 * needed is in node:readline.
 */

const cyan = (value: string): string => styleText("cyan", value);
const dim = (value: string): string => styleText("dim", value);
const green = (value: string): string => styleText("green", value);

export class PromptCancelledError extends Error {
  constructor() {
    super("cancelled");
    this.name = "PromptCancelledError";
  }
}

export interface Choice<T extends string> {
  value: T;
  label: string;
  hint?: string;
}

/** Ask for a line of text, returning `fallback` when the answer is empty. */
export async function text(question: string, fallback: string): Promise<string> {
  const rl = readline.promises.createInterface({ input: stdin, output: stdout });
  try {
    const answer = await rl.question(`${green("?")} ${question} ${dim(`(${fallback})`)} `);
    return answer.trim() || fallback;
  } finally {
    rl.close();
  }
}

/**
 * Arrow-key selection.
 *
 * Falls back to a numbered list when stdin is not a TTY — piped input and CI
 * would otherwise hang forever waiting for a keypress that cannot come.
 */
export async function select<T extends string>(
  question: string,
  choices: Array<Choice<T>>,
): Promise<T> {
  if (!stdin.isTTY) return numbered(question, choices);

  const width = Math.max(...choices.map((choice) => choice.label.length));
  let active = 0;

  const paint = (): void => {
    const lines = choices.map((choice, index) => {
      const label = index === active ? cyan(`❯ ${choice.label}`) : `  ${choice.label}`;
      const pad = " ".repeat(width - choice.label.length);
      return choice.hint ? `${label}${pad}  ${dim(choice.hint)}` : label;
    });
    stdout.write(`${green("?")} ${question}\n${lines.join("\n")}\n`);
  };

  const clear = (): void => {
    readline.moveCursor(stdout, 0, -(choices.length + 1));
    readline.clearScreenDown(stdout);
  };

  readline.emitKeypressEvents(stdin);
  stdin.setRawMode(true);
  // Closing the readline interface of an earlier prompt leaves stdin paused, and
  // attaching a `keypress` listener does not resume it. Without this the stream
  // never flows, the event loop runs dry, and node exits in the middle of the
  // await with "Detected unsettled top-level await".
  stdin.resume();
  stdout.write("\u001B[?25l"); // hide the cursor
  paint();

  try {
    return await new Promise<T>((resolve, reject) => {
      const onKeypress = (_: string, key: { name?: string; ctrl?: boolean }): void => {
        if (key.ctrl && key.name === "c") {
          stdin.off("keypress", onKeypress);
          reject(new PromptCancelledError());
          return;
        }
        if (key.name === "up" || key.name === "k")
          active = (active - 1 + choices.length) % choices.length;
        else if (key.name === "down" || key.name === "j") active = (active + 1) % choices.length;
        else if (key.name === "return") {
          stdin.off("keypress", onKeypress);
          resolve(choices[active]!.value);
          return;
        } else return;

        clear();
        paint();
      };
      stdin.on("keypress", onKeypress);
    });
  } finally {
    stdout.write("\u001B[?25h"); // show it again
    stdin.setRawMode(false);
    stdin.pause();
    clear();
  }
}

async function numbered<T extends string>(question: string, choices: Array<Choice<T>>): Promise<T> {
  const rl = readline.promises.createInterface({ input: stdin, output: stdout });
  try {
    const list = choices.map((choice, index) => `  ${index + 1}) ${choice.label}`).join("\n");
    const answer = await rl.question(`${green("?")} ${question}\n${list}\n  ${dim("[1]")} `);
    const index = Number.parseInt(answer.trim() || "1", 10) - 1;
    return (choices[index] ?? choices[0]!).value;
  } finally {
    rl.close();
  }
}
