import win32com
import win32com.client
import os
import time


class CANalyzer:
    def __init__(self):
        self.application = None
        self.application = win32com.client.DispatchEx("CANalyzer.Application")
        if self.application is None:
            raise RuntimeError("Start CANalyzer Application is failed, unable to open simulation")
        else:
            self.ver = self.application.Version
            print('CANalyzer version ',
                  self.ver.major, '.',
                  self.ver.minor, '.',
                  self.ver.Build)
            self.Measurement = self.application.Measurement.Running
            if self.Measurement:
                self.application.Measurement.Stop()
            print('Active Measurement status is:', self.Measurement)

    def getCurrentConfigurationPath(self):
        if (self.application != None):
            conf_obj = self.application.Configuration
            configurationPath = conf_obj.FullName
            print('Configuration: ', configurationPath)
            return str(configurationPath)
        else:
            raise RuntimeError("Can't find started CANalyzer")

    def loadExistedConfiguration(self, cfgname):
        if (self.application != None):
            if os.path.isfile(cfgname) and (os.path.splitext(cfgname)[1] == ".cfg"):
                self.application.Open(cfgname)
            else:
                raise RuntimeError("Can't find CANalyzer cfg file")

    # TODO: find way to assign dbc separately from cfg file. now just return assigned channel from dbc
    def loadNewDBCtoConfiguration(self):
        dbSetup = self.application.Configuration.GeneralSetup.DatabaseSetup
        db = dbSetup.Databases(1)
        channel = db.Channel
        print("Assigned channel: ", channel)

    def getBusStatistic(self, busType, channel):
        if (self.application != None):
            busStatistics = self.application.MeasurementSetup.BusStatistics.BusStatistic(busType, channel)
            return str(busStatistics)
        else:
            raise RuntimeError("Can't find started CANalyzer")
    def startMeasurement(self):
        retry = 0
        retry_counter = 5
        while not self.application.Measurement.Running and (retry < retry_counter):
            self.application.Measurement.Start()
            time.sleep(1)
            retry += 1
        if (retry == retry_counter):
            raise RuntimeWarning("CANalyzer start measurement failed, Please Check Connection!")

    def stopMeasurement(self):
        if self.application.Measurement.Running:
            self.application.Measurement.Stop()
        else:
            pass

    def getSignalValue(self, channel, sig_message, sig_name):
        if self.application is not None:
            retval = self.application.GetBus("CAN").GetSignal(channel, sig_message, sig_name)
            return retval.Value
        else:
            raise RuntimeWarning('CANalyzer application is not opened')

    def getSignalState(self, channel, sig_message, sig_name):
        if self.application is not None:
            retval = self.application.GetBus("CAN").GetSignal(channel, sig_message, sig_name)
            return retval.State
        else:
            raise RuntimeWarning('CANalyzer application is not opened')

    def getStateIsSignalOnline(self, channel, sig_message, sig_name):
        if self.application is not None:
            retval = self.application.GetBus("CAN").GetSignal(channel, sig_message, sig_name)
            return retval.IsOnline
        else:
            raise RuntimeWarning('CANalyzer application is not opened')

    def setSignalActiveValue(self, channel, sig_message, sig_name, sig_value):
        if self.application is not None:
            canSignal = self.application.GetBus("CAN").GetSignal(channel, sig_message, sig_name)
            canSignal.Value = sig_value
        else:
            raise RuntimeWarning('CANalyzer application is not opened')

if __name__ == '__main__':
    app = CANalyzer()
    conf = os.path.basename("tmp.cfg")
    app.loadExistedConfiguration(conf)
    app.getCurrentConfigurationPath()
    app.loadNewDBCtoConfiguration()
    app.startMeasurement()

    #app.getSignalValue(1, "FrontWasherControl", "Wipers")

    # sys_value = self.sys_namespace.Variables
    # result = self.sys_value(self.var)

    # result = app.sys_value(self.var)


