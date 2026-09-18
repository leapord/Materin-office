/**
 * Shared view fragments: banner strip and status line.
 */

export interface BannerAction {
	label: string;
	onClick: () => void;
}

export interface Banner {
	setText(text: string): void;
	destroy(): void;
}

export function showBanner(
	host: HTMLElement,
	text: string,
	tone: "info" | "warning" | "error" = "info",
	actions: BannerAction[] = [],
): Banner {
	const el = host.createDiv(`materin-office-banner materin-office-banner-${tone}`);
	el.createSpan({ text });
	const bar = el.createDiv("materin-office-banner-actions");
	for (const action of actions) {
		bar.createEl("button", {
			text: action.label,
			cls: "mod-muted",
		}).addEventListener("click", action.onClick);
	}
	return {
		setText(next: string) {
			el.firstElementChild?.setText(next);
		},
		destroy() {
			el.remove();
		},
	};
}

export interface StatusLine {
	setText(text: string): void;
}

export function createStatusLine(host: HTMLElement): StatusLine {
	const el = host.createDiv("materin-office-status");
	return {
		setText(text: string) {
			el.setText(text);
		},
	};
}
