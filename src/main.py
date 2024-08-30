try:
    from src import colvir
except ModuleNotFoundError as error:
    raise error


def main() -> None:
    colvir.run()


if __name__ == "__main__":
    main()
