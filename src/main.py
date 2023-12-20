'''
Main function handling the frontend and backend 
'''


def main():
    # Init server-side
    from server import EmesDispatch
    emes = EmesDispatch()

    # Init client-side
    from ui import run_ui
    run_ui(emes)


if __name__ == '__main__':
    main()
