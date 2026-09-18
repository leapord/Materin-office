/** Binary/encoding helpers shared by all format engines (obsidian-free). */

export function toBytes(data: ArrayBuffer): Uint8Array {
	return new Uint8Array(data);
}

/** Copies the bytes into a fresh ArrayBuffer (avoids ArrayBufferLike typing). */
export function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	const copy = new ArrayBuffer(bytes.byteLength);
	new Uint8Array(copy).set(bytes);
	return copy;
}

export function decodeUtf8(bytes: Uint8Array): string {
	// ignoreBOM: true means "keep a leading BOM in the output" (do not strip);
	// BOM preservation is handled explicitly by XmlPart.
	return new TextDecoder("utf-8", { ignoreBOM: true }).decode(bytes);
}

export function encodeUtf8(text: string): Uint8Array {
	return new TextEncoder().encode(text);
}

export function concatBytes(chunks: readonly Uint8Array[]): Uint8Array {
	const total = chunks.reduce((n, chunk) => n + chunk.length, 0);
	const out = new Uint8Array(total);
	let offset = 0;
	for (const chunk of chunks) {
		out.set(chunk, offset);
		offset += chunk.length;
	}
	return out;
}

export const UTF8_BOM = new Uint8Array([0xef, 0xbb, 0xbf]);

export function startsWithUtf8Bom(bytes: Uint8Array): boolean {
	return (
		bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
	);
}
