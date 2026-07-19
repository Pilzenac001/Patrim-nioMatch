import sys
from interface import AppPatrimonioMatch

def main():
    try:
        app = AppPatrimonioMatch()
        app.mainloop()
    except Exception as e:
        print(f"Ocorreu um erro ao iniciar a aplicação: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
