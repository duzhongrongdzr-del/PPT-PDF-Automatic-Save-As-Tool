"""Entry point; keep GUI imports after freeze_support for child processes."""
import multiprocessing
import sys


def main():
    multiprocessing.freeze_support()
    if "--self-test" in sys.argv:
        from zhiting.selftest import run
        run(sys.argv[-1])
        return
    if sys.platform != "win32":
        raise SystemExit("此程序需要 Windows 10 / 11。")
    from zhiting.app import run
    run(background="--background" in sys.argv)


if __name__ == "__main__":
    main()
