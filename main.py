"""Entry point; keep GUI imports after freeze_support for child processes."""
import multiprocessing
import sys


def main():
    multiprocessing.freeze_support()
    if "--self-test" in sys.argv:
        try:
            from zhiting.selftest import run
            run(sys.argv[-1])
        except Exception:
            import json
            from pathlib import Path
            import traceback
            Path(sys.argv[-1]).write_text(json.dumps({"ok": False, "error": traceback.format_exc()}), encoding="utf-8")
            raise SystemExit(1)
        return
    if sys.platform != "win32":
        raise SystemExit("此程序需要 Windows 10 / 11。")
    from zhiting.app import run
    run(background="--background" in sys.argv)


if __name__ == "__main__":
    main()
