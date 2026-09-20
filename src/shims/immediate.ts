/*
 * Build-time replacement for the `immediate` package (pulled in via jszip → lie).
 *
 * The original falls back to <script> onreadystatechange injection and
 * `new Function` string coercion on ancient browsers. Obsidian runs on modern
 * Chromium where MutationObserver always exists, so those branches are dead
 * code — but Obsidian's plugin static review flags their presence in the
 * bundle. This shim keeps the exact same task-queue/drain semantics and only
 * the safe schedulers.
 */

type Task = () => void;

const MutationObserverCtor =
	window.MutationObserver ?? (window as unknown as Record<string, typeof MutationObserver>).WebKitMutationObserver;

type ScheduleDrain = () => void;
let scheduleDrain: ScheduleDrain;

if (typeof MutationObserverCtor === "function") {
	let called = 0;
	const element = document.createTextNode("");
	const observer = new MutationObserverCtor(drain);
	observer.observe(element, { characterData: true });
	scheduleDrain = () => {
		element.data = (called = ++called % 2) as unknown as string;
	};
} else if (typeof window.MessageChannel === "function") {
	const channel = new window.MessageChannel();
	channel.port1.onmessage = drain;
	scheduleDrain = () => {
		channel.port2.postMessage(0);
	};
} else {
	scheduleDrain = () => {
		window.setTimeout(drain, 0);
	};
}

let draining = false;
const queue: Task[] = [];

function drain(): void {
	draining = true;
	let oldQueue: Task[];
	let len = queue.length;
	while (len) {
		oldQueue = queue;
		queue.length = 0;
		for (const task of oldQueue) {
			task();
		}
		len = queue.length;
	}
	draining = false;
}

const immediate = (task: Task): void => {
	if (queue.push(task) === 1 && !draining) {
		scheduleDrain();
	}
};

export default immediate;
