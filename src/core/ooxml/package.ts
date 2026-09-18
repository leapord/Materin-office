import JSZip from "jszip";
import { CorruptFileError, MissingPartError } from "../errors";

interface Part {
	original: Uint8Array;
	edited: Uint8Array | null;
}

/**
 * An OOXML zip package with byte-preserving semantics: untouched parts keep
 * their original bytes verbatim, only explicitly edited parts are replaced.
 * (Rebuilding the zip container changes container-level bytes — OOXML has no
 * ordering requirement, so that is safe; Word/PowerPoint only care about part
 * contents.)
 */
export class OoxmlPackage {
	private readonly parts = new Map<string, Part>();

	private constructor() {}

	static async load(data: ArrayBuffer): Promise<OoxmlPackage> {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(data);
		} catch (cause) {
			throw new CorruptFileError(`无法解析 OOXML 包（zip 结构损坏）：${String(cause)}`);
		}
		const pkg = new OoxmlPackage();
		const entries = Object.values(zip.files).filter((entry) => !entry.dir);
		await Promise.all(
			entries.map(async (entry) => {
				const bytes = await entry.async("uint8array");
				pkg.parts.set(entry.name, { original: bytes, edited: null });
			}),
		);
		return pkg;
	}

	hasPart(path: string): boolean {
		return this.parts.has(path);
	}

	partPaths(): string[] {
		return [...this.parts.keys()];
	}

	/** Current bytes of a part: edited bytes if present, else the original. */
	getPartBytes(path: string): Uint8Array {
		const part = this.parts.get(path);
		if (!part) {
			throw new MissingPartError(path);
		}
		return part.edited ?? part.original;
	}

	setPartBytes(path: string, bytes: Uint8Array): void {
		const part = this.parts.get(path);
		if (!part) {
			throw new MissingPartError(path);
		}
		part.edited = bytes;
	}

	isEdited(path: string): boolean {
		const part = this.parts.get(path);
		return part?.edited !== null && part?.edited !== undefined;
	}

	editedPartPaths(): string[] {
		return [...this.parts.entries()]
			.filter(([, part]) => part.edited !== null)
			.map(([path]) => path);
	}

	/** Rebuilds the zip: edited parts from their new bytes, all others original. */
	async serialize(): Promise<Uint8Array> {
		const zip = new JSZip();
		for (const [path, part] of this.parts) {
			zip.file(path, part.edited ?? part.original);
		}
		return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
	}
}
