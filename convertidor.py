from cx_Freeze import setup, Executable

setup(name = "Robot Trading",
	version = "9.10",
	description = "Par de monedas USD/JPY",
	executables = [Executable("mind_control_v9_10.py")],)