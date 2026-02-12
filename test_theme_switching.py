import TKinterModernThemes.examples.allwidgets as allwidgets
import TKinterModernThemes.examples.layoutdemo as layoutdemo

def test_every_theme():
    for mode in ['dark', 'light']:
        for theme in ['azure', 'forest', 'sun-valley']:
            print("Running test with mode: ", mode, " and theme: ", theme)
            allwidgets.App(theme, mode)

def test_layout_engine(capture_stdout):
    for mode in ['dark', 'light']:
        for theme in ['azure', 'forest', 'sun-valley']:
            print("Running test with mode: ", mode, " and theme: ", theme)
            capture_stdout['stdout'] = ""
            layoutdemo.App(theme, mode)
            assert capture_stdout['stdout'] == open('TKinterModernThemes/examples/layoutdemooutput.txt').read()

def test_load():
    import TKinterModernThemes #forces a basic test of imports
    TKinterModernThemes.firstWindow = False #make sure actually imports
    
if __name__ == "__main__":
    test_every_theme()
    test_layout_engine({'stdout': ''})
    test_load()