function testCOMApi(){
	
	//consuming Win32SleepAwakeTime
	console.log('Last awake and sleep time info:');
	var win32SleepAwakeTime = new ActiveXObject('PowerStateManagmentCOM.Win32SleepAwakeTime');
	var dtLastAwakeTime = win32SleepAwakeTime.GetLastAwakeTime();
	var dtLastSleepTime = win32SleepAwakeTime.GetLastSleepTime();
	console.log('Last awake time: ' + dtLastAwakeTime);
	console.log('Last sleep time: ' + dtLastSleepTime);
	console.log('');
	console.log('');
	
	//consuming Win32SystemBatteryState
	console.log('System battery state info:');
	console.log('---------------------------------------------------------------');
	var win32SystemBatteryState = new ActiveXObject("PowerStateManagmentCOM.Win32SystemBatteryState");
	var batteryState = win32SystemBatteryState.GetSystemBatteryState();
	console.log('AcOnLine: ' + batteryState.AcOnLine);
	console.log('Battery Present: ' + batteryState.BatteryPresent);
	console.log('Charging: ' + batteryState.Charging);
	console.log('Discharging: ' + batteryState.Discharging);
	console.log('');
	console.log('');
	
	//usage of Win32SystemPowerInformation
	console.log('System Power Information:');
	console.log('---------------------------------------------------------------');
	var win32SystemPowerInformation = new ActiveXObject("PowerStateManagmentCOM.Win32SystemPowerInformation");
	var powerInfo = win32SystemPowerInformation.GetSystemPowerInformation();
	console.log('Cooling mode: ' + powerInfo.CoolingMode);
	console.log('Idleness: ' + powerInfo.Idleness);
	console.log('MaxIdlenessAllowed: ' + powerInfo.MaxIdlenessAllowed);
	console.log('Time remaining: ' + powerInfo.TimeRemaining);
}