/**
 * jsdom polyfill for the Obsidian DOM extensions the view layer relies on
 * (createDiv/createEl/createSpan/addClass/empty/setText/instanceOf/
 * setCssProps). Lets KindUi components render under vitest without Obsidian.
 *
 * The polyfill defines the Obsidian helpers themselves, so it must create
 * elements with the raw DOM API — eslint's prefer-create-el is a plugin-code
 * rule that cannot apply to the polyfill that provides those helpers.
 */

interface ObsidianDomExtensions {
	createDiv(
		spec?: string | { text?: string; cls?: string },
	): HTMLElement;
	createEl(
		tag: string,
		spec?: string | { text?: string; cls?: string },
	): HTMLElement;
	createSpan(spec?: { text?: string; cls?: string }): HTMLElement;
	addClass(cls: string): void;
	empty(): void;
	setText(text: string): void;
	setCssProps(props: Record<string, string>): void;
	instanceOf(ctor: abstract new () => HTMLElement): boolean;
}

/** Installs the polyfill on a window's HTMLElement prototype. */
export function installObsidianDomPolyfill(win: {
	HTMLElement: typeof HTMLElement;
}): void {
	// Bound reference: this polyfill IS the provider of createEl/createDiv —
	// the raw DOM API is the only correct primitive at this layer. The
	// structural type sidesteps the DOM-lib's deprecated string overload.
	const docSource: { createElement(tag: string): HTMLElement } = document;
	const createRaw = docSource.createElement.bind(docSource);
	const proto = win.HTMLElement.prototype as unknown as ObsidianDomExtensions;
	proto.createDiv = function (
		this: HTMLElement,
		spec?: string | { text?: string; cls?: string },
	): HTMLElement {
		const div = createRaw("div");
		applySpec(div, spec);
		this.appendChild(div);
		return div;
	};
	proto.createEl = function (
		this: HTMLElement,
		tag: string,
		spec?: string | { text?: string; cls?: string },
	): HTMLElement {
		const el = createRaw(tag);
		applySpec(el, spec);
		this.appendChild(el);
		return el;
	};
	proto.createSpan = function (
		this: HTMLElement,
		spec?: { text?: string; cls?: string },
	): HTMLElement {
		const span = createRaw("span");
		applySpec(span, spec);
		this.appendChild(span);
		return span;
	};
	proto.addClass = function (this: HTMLElement, cls: string): void {
		this.classList.add(cls);
	};
	proto.empty = function (this: HTMLElement): void {
		this.textContent = "";
	};
	proto.setText = function (this: HTMLElement, text: string): void {
		this.textContent = text;
	};
	proto.setCssProps = function (
		this: HTMLElement,
		props: Record<string, string>,
	): void {
		for (const [name, value] of Object.entries(props)) {
			this.style.setProperty(name, value);
		}
	};
	proto.instanceOf = function (
		this: HTMLElement,
		ctor: abstract new () => HTMLElement,
	): boolean {
		// `instanceof` here would recurse into the method being implemented;
		// the prototype-chain check is the canonical cross-window-safe form.
		return Object.prototype.isPrototypeOf.call(ctor.prototype, this);
	};
}

/** Applies {text, cls} specs the way Obsidian's helpers do. */
function applySpec(
	el: HTMLElement,
	spec?: string | { text?: string; cls?: string },
): void {
	const cls = typeof spec === "string" ? spec : spec?.cls;
	if (cls) {
		el.className = cls;
	}
	if (typeof spec === "object" && spec?.text !== undefined) {
		el.textContent = spec.text;
	}
}

/**
 * Prepares the vitest jsdom environment's own document as a render host.
 * Uses the GLOBAL window (not a fresh JSDOM) so `instanceof HTMLElement`
 * checks inside view code see the same class the environment provides.
 */
export function createHost(): HTMLElement {
	installObsidianDomPolyfill(window);
	return document.body;
}
