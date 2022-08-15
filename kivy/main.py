from kivy.app import App
from kivy.lang import Builder
from kivy.uix.button import Button

GUI = Builder.load_file('tela.kv')


class NavarroApp(App):
    def build(self):
        return GUI


NavarroApp().run()
