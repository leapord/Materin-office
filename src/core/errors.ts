/** Base class for all Office-processing errors the plugin raises deliberately. */
export class OfficeFileError extends Error {}

/** The file is not a valid zip / OOXML package (truncated sync writes, wrong format). */
export class CorruptFileError extends OfficeFileError {}

/** The file is a recognized Office kind we cannot process (e.g. legacy .ppt). */
export class UnsupportedFormatError extends OfficeFileError {}

/** A required OOXML part is missing from the package. */
export class MissingPartError extends OfficeFileError {
	constructor(readonly partPath: string) {
		super(`OOXML 包缺少必需部件：${partPath}`);
	}
}
