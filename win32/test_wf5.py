import sys
import os
sys.path.insert(0, r'C:\hwp_xml')
sys.path.insert(0, r'C:\hwp_xml\win32')
os.chdir(r'C:\hwp_xml')

from workflow.workflow5_integrated import Workflow5

workflow5 = Workflow5()
workflow5.run(r'C:\hwp_xml\test.hwp')
