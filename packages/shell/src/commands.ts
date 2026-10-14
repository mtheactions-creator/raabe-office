// The command registry: one place for every command's label, icon, shortcut and state.
// Ribbon buttons, the Quick Access Toolbar, keyboard shortcuts and the search box all
// execute commands by id. A command without `run` is planned but not built yet; it shows
// greyed out (agreed for development builds).

export interface Command {
  id: string;
  /** German label as in Word, e.g. "Fett". */
  label: string;
  icon?: string;
  /** Longer tooltip text; the label is used when missing. */
  tip?: string;
  /** Keyboard shortcut shown in tooltips, e.g. "Strg+S". */
  shortcut?: string;
  run?: () => void | Promise<void>;
  enabled?: () => boolean;
  checked?: () => boolean;
  /** Hidden from the search box (e.g. internal commands). */
  hidden?: boolean;
  /** Extra search terms ("Inhaltsverzeichnis" for "toc"). */
  keywords?: string[];
  /** Deliberately postponed (SCOPE, D-041): shown greyed out and marked in the tooltip. */
  deferred?: boolean;
}

type Listener = () => void;

export class Commands {
  private map = new Map<string, Command>();
  private listeners = new Set<Listener>();
  private scheduled = false;

  register(...commands: Command[]): void {
    for (const c of commands) this.map.set(c.id, c);
  }

  /** Adds behaviour to a registered command (the app wires up what is implemented). */
  implement(id: string, impl: Pick<Command, "run" | "enabled" | "checked">): void {
    const c = this.get(id);
    Object.assign(c, impl);
  }

  get(id: string): Command {
    const c = this.map.get(id);
    if (!c) throw new Error(`unknown command ${id}`);
    return c;
  }

  has(id: string): boolean {
    return this.map.has(id);
  }

  all(): Command[] {
    return [...this.map.values()];
  }

  isImplemented(id: string): boolean {
    return !!this.map.get(id)?.run;
  }

  isEnabled(id: string): boolean {
    const c = this.map.get(id);
    return !!c?.run && (c.enabled ? c.enabled() : true);
  }

  isChecked(id: string): boolean {
    return !!this.map.get(id)?.checked?.();
  }

  async run(id: string): Promise<void> {
    if (!this.isEnabled(id)) return;
    await this.get(id).run!();
    this.refresh();
  }

  /** Called whenever state may have changed; UI updates are batched into one frame. */
  refresh(): void {
    if (this.scheduled) return;
    this.scheduled = true;
    requestAnimationFrame(() => {
      this.scheduled = false;
      for (const l of this.listeners) l();
    });
  }

  onRefresh(l: Listener): void {
    this.listeners.add(l);
  }

  /** `available` overrides whether the command works (a menu button works when its menu does). */
  tooltip(id: string, available?: boolean): string {
    const c = this.get(id);
    let t = c.tip ?? c.label;
    if (c.shortcut) t += ` (${c.shortcut})`;
    if (c.deferred) t += "\nZurückgestellt: bewusst auf eine spätere Version verschoben";
    else if (!(available ?? !!c.run)) t += "\nNoch nicht verfügbar";
    return t;
  }
}
