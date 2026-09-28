/** Unsaved-changes flag behind `element.dirty`; reports only real transitions. */
export class DirtyState {
	private value = false;

	constructor(private readonly changed: (dirty: boolean) => void) {}

	get dirty(): boolean {
		return this.value;
	}

	set(dirty: boolean): void {
		if (this.value === dirty) return;
		this.value = dirty;
		this.changed(dirty);
	}
}
