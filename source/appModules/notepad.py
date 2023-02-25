# A part of NonVisual Desktop Access (NVDA)
# Copyright (C) 2022 NV Access Limited, Joseph Lee
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

"""App module for Windows Notepad.
While this app module also covers older Notepad releases,
this module provides workarounds for Windows 11 Notepad."""

from NVDAObjects.IAccessible import IAccessible
from logHandler import log
from scriptHandler import script

import api
import appModuleHandler
import controlTypes
import requests
import ui


class AppModule(appModuleHandler.AppModule):

	def _get_statusBar(self):
		"""Retrieves Windows 11 Notepad status bar.
		In Windows 10 and earlier, status bar can be obtained by looking at the bottom of the screen.
		Windows 11 Notepad uses Windows 11 UI design (top-level window is labeled "DesktopWindowXamlSource",
		therefore status bar cannot be obtained by position alone.
		If visible, a child of the foreground window hosts the status bar elements.
		Status bar child position must be checked whenever Notepad is updated on stable Windows 11 releases
		as Notepad is updated through Microsoft Store as opposed to tied to specific Windows releases.
		L{api.getStatusBar} will resort to position lookup if C{NotImplementedError} is raised.
		"""
		# #13688: Notepad 11 uses Windows 11 user interface, therefore status bar is harder to obtain.
		# This does not affect earlier versions.
		notepadVersion = int(self.productVersion.split(".")[0])
		if notepadVersion < 11:
			raise NotImplementedError()
		# And no, status bar is shown when editing documents.
		# Thankfully, of all the UIA objects encountered, document window has a unique window class name.
		if api.getFocusObject().windowClassName != "RichEditD2DPT":
			raise NotImplementedError()
		# Look for a specific child as some children report the same UIA properties such as class name.
		# Make sure to look for a foreground UIA element which hosts status bar content if visible.
		notepadStatusBarIndex = 7
		statusBar = api.getForegroundObject().children[notepadStatusBarIndex].firstChild
		# No location for a disabled status bar i.e. location is 0 (x, y, width, height).
		if not any(statusBar.location):
			raise NotImplementedError()
		return statusBar

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if obj.windowClassName == "Edit" and obj.role == controlTypes.Role.EDITABLETEXT:
			clsList.insert(0, EnhancedEditField)

class EnhancedEditField(IAccessible):

	@script(gesture="kb:NVDA+l")
	def script_reportTypos(self, gesture):
		ui.message(f"原文是: {self.value}")
		text_corrected = self._get_openai_completion_response(self.value)
		self._report_typos(self.value,
				   text_corrected)

	def _report_typos(self, text_original, text_corrected):
		text_original_split = text_original.split('\r\n')
		text_corrected_split = text_corrected.split('\n')
		error_count = 0

		ui.message(f"分析結果如下:")

		for row in range(len(text_original_split)):
			for col in range(len(text_original_split[row])):
				if text_original_split[row][col] != text_corrected_split[row][col]:
					ui.message(f"row {row + 1}, column {col + 1}: {text_original_split[row][col]} 應改成 {text_corrected_split[row][col]}")
					error_count += 1

		ui.message(f"共有錯字{error_count}個")

	def _get_openai_completion_response(self, prompt_text):
		prompt_augmented = f'改錯字\n\n題目:{prompt_text}\n\n答案:'

		API_KEY = 'OPENAI_API_KEY'
		url =  "https://api.openai.com/v1/completions"
		headers = {"Authorization": f"Bearer {API_KEY}"}
		data = {'model': 'text-davinci-003',
			'prompt': prompt_augmented,
			'max_tokens': 60,
			'temperature': 0,
			}

		response = requests.post(url, headers=headers, json=data).json()

		return response['choices'][0]['text']
