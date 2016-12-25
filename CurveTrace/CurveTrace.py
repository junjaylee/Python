################################################
#!/usr/bin/env python
# coding: utf-8
# V1 : Initialize Version
#	Record Measurment Voltage and Current and generate chart
################################################
import sys
import ConfigParser
import logging
import pandas as pd
import visa
import time
import math
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from pylab import savefig
from pyvisa.errors import (Error, VisaIOError)

Tool_Ver = 1.2
logging.basicConfig(filename='CurveTrace.log',level=logging.DEBUG)
logging.info('Tool Version = %f',Tool_Ver)
config = ConfigParser.ConfigParser({'PPS_Addr':'GPIB0::5','Delay':0.5,'CURR_Limit':1,'Grid':'false', 'SkipInstrument':'false'})
config.read('Config.ini')
PPS_Addr = config.get('Configure', 'PPS_Addr')
Delay = config.getfloat('Configure','Delay')
bGrid = config.getboolean('Configure','Grid')
CURR_Limit = config.getfloat('Configure','CURR_Limit')
bSkipInstrument = config.getboolean('Configure','SkipInstrument')
logging.info('PPS_Addr = %s,Delay = %f',PPS_Addr,Delay)
print 'CurveTrace Tool Version {0}'.format(Tool_Ver)
try:
	rm = visa.ResourceManager()
	if (bSkipInstrument == False):
		PPS_inst = rm.get_instrument(PPS_Addr)
		PPS_inst.write("*RST")
		PPS_inst.write("*CLS")
		logging.info('IDN = %s',PPS_inst.ask("*IDN?"))
		PPS_inst.write('CURRent {0}'.format(CURR_Limit))
		PPS_inst.write("OUTPUT ON")
except VisaIOError as e:
	logging.info('PPS_Addr = %s (%d, %s)',PPS_Addr, e.error_code, e.description)
	sys.exit(0)	
arg_total = len(sys.argv)
if (arg_total == 5):
	Chart_Title = sys.argv[1]
	Start_Voltage = float(sys.argv[2])
	End_Voltage = float(sys.argv[3])
	Step_Voltage = float(sys.argv[4])
else:
	Chart_Title = raw_input('Please input Chart Title ===>')
	Start_Voltage = float(raw_input('Please input start of Voltage ===>'))
	End_Voltage = float(raw_input('Please input end of Voltage ===>'))
	Step_Voltage = float(raw_input('Please input measurment step of Voltage ===>'))
logging.info('Chart Title = %s,Start Voltage = %f,End Voltage = %f,Step Voltage = %f',Chart_Title,Start_Voltage,End_Voltage,Step_Voltage)
count = Start_Voltage
Voltage_Data = pd.Series()
Current_Data = pd.Series()
if (Start_Voltage < 0):
	raw_input('Please probe GND to Positive net')
bPrompt = False
while (round(count,5) <= round(End_Voltage,5)):
	Meas_Data = 0
	Data = count
	if (bSkipInstrument == False):
		if (count < 0):
			PPS_inst.write('VOLT {:.2}'.format(math.fabs(count)))
		else:
			PPS_inst.write('VOLT {:.2}'.format(count))
		time.sleep(Delay)
		Meas_Data = PPS_inst.ask("MEAS:CURR?")
		Data = float(Meas_Data.strip()) * 1000
		if (count < 0):
			Data *= -1
	Voltage_Data = Voltage_Data.append(pd.Series([count]))
	Current_Data = Current_Data.append(pd.Series([Data]))
	count = count + Step_Voltage
	if (Start_Voltage < 0 and count > 0 and bPrompt == False):
		raw_input('Please probe Positive net to GND')
		bPrompt = True
if (bSkipInstrument == False):
	PPS_inst.write("OUTPUT OFF")
DataFrame = pd.DataFrame({ 'Voltage' : Voltage_Data,'Current' : Current_Data})
print DataFrame
logging.info('Measurment DataFrame = %s',DataFrame)
fig = plt.figure()
minorLocator = MultipleLocator(Step_Voltage)
plt.plot(DataFrame['Voltage'], DataFrame['Current'])
plt.grid(bGrid)
ax = plt.gca()
ax.xaxis.set_minor_locator(minorLocator)
plt.axhline(y=0, color='k', linestyle='-.')
plt.axvline(x=0, color='k', linestyle='-.')
plt.title(Chart_Title)
plt.ylabel('mA')
plt.xlabel('V')
plt.tight_layout()
savefig('{0}_Chart.png'.format(Chart_Title))
plt.close('all')
logging.info('Finished')
print 'Finished'
