import time
import os
import sys
import argparse
import psutil
from pywinauto.application import Application
from pywinauto import timings
from pywinauto.keyboard import send_keys


def type_keys(string, element):
    """Type a string char by char to Element window"""
    for char in string:
        element.type_keys(char)

def main():   
	# Parse arguments from cmd
	parser = argparse.ArgumentParser()
	parser.add_argument("workbook", help = "Path to .pbix file")
	parser.add_argument("--workspace", help = "name of online Power BI service work space to publish in", default = "My workspace")
	parser.add_argument("--refresh-timeout", help = "refresh timeout", default = 30000, type = int)
	parser.add_argument("--publish", dest='publish', help="don't publish, just save", default = True, action = 'store_false' )
	parser.add_argument("--init-wait", help = "initial wait time on startup", default = 25, type = int)
	args = parser.parse_args()

	timings.after_clickinput_wait = 1
	WORKBOOK = args.workbook
	WORKSPACE = args.workspace
	INIT_WAIT = args.init_wait
	REFRESH_TIMEOUT = args.refresh_timeout
	FILE_NAME = args.workbook.split("\\")[-1]

	# Kill running PBI
	PROCNAME = "PBIDesktop.exe"
	for proc in psutil.process_iter():
		# check whether the process name matches
		if proc.name() == PROCNAME:
			proc.kill()
	time.sleep(3)

	# Start PBI and open the workbook
	print(f"Now refreshing: {WORKBOOK}")
	print("Starting Power BI")
	os.system('start "" "' + WORKBOOK + '"')
	print("Waiting ",INIT_WAIT,"sec")
	time.sleep(INIT_WAIT)

	# Connect pywinauto
	print("Identifying Power BI window")
	app = Application(backend = 'uia').connect(path = PROCNAME)
	win = app[FILE_NAME]
	print("Window identified")
	time.sleep(5)
	win.set_focus()
	win.wait("enabled", timeout=300)

	# Refresh
	print("Refreshing")
	send_keys("{VK_MENU} h r {DOWN} {ENTER}", pause=0.2)
	time.sleep(5)
	print("Waiting for refresh end (timeout in ", REFRESH_TIMEOUT,"sec)")
	win.wait("enabled", timeout = REFRESH_TIMEOUT)

	# Save
	print("Saving")
	send_keys('^s')
	win.wait("enabled", timeout = REFRESH_TIMEOUT)
	time.sleep(5)

	# Publish
	if args.publish:
		print("Publish")
		send_keys("{VK_MENU} h p ", pause=0.2)
		time.sleep(2)
		publish_dialog = win.child_window(title="Publish to Power BI", auto_id="mat-mdc-dialog-0", control_type="Group")
		if publish_dialog.child_window(title = WORKSPACE).exists():
			publish_dialog.child_window(title = WORKSPACE).click_input()
			publish_dialog.child_window(title="Select", auto_id="okButton", control_type="Button").click_input()
			try:
				win.child_window(title="Replace", auto_id="okButton", control_type="Button").wait('visible', timeout = 10)
			except Exception:
				pass
			if win.child_window(title="Replace", auto_id="okButton", control_type="Button").exists():
				win.child_window(title="Replace", auto_id="okButton", control_type="Button").click_input()
			win.child_window(title="Got it", control_type="Button").wait('visible', timeout = REFRESH_TIMEOUT)
			win.child_window(title="Got it", control_type="Button").click_input()
		else:
			print("Workspace not found")

	#Close
	# Workaround - window is "enabled" even if the save did not finish (corrupts files)
	#attempt to close several times
	while any(proc.name() == PROCNAME for proc in psutil.process_iter()):
		print("Attempting to close")
		win.close()
		time.sleep(5)

	print("Refresh successful")

		
if __name__ == '__main__':
	try:
		main()
	except Exception as e:
		print(e)
		sys.exit(1)








