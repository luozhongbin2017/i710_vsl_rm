# This file defines the class of ramp meter.

class RampMeter(object):
    """
    LINK_GROUP: (int) the link group index of the section that the ramp meter is conneted to
    DC: (VISSIM Object) the VISSIM object of data collection on the ramp
    SC: (VISSIM Object) the VISSIM object of the signal controller associated with the ramp meter
    QC: (VISSIM Object) the VISSIM object of the queue counter associated with the ramp meter
    QLENGTH: (double) reference queue length on the ramp
    RHOR: (double) reference density of the connected highway section
    redTime: expected red light time 
    greenTime: expected green light time
    RedTimer: (double) timer for the red light
    GreenTimer: (double) timer for the greent light
    """
    def __init__(self, LINK_GROUP, DC, SC, QC, QLENGTH, RHOR):
        #super(RampMeter, self).__init__()
        self.LINK_GROUP = LINK_GROUP
        self.DC = DC
        self.SC = SC
        self.QC = QC
        self.SG = self.SC.SignalGroups.GetSignalGroupByNumber(1)
        self.Det = self.SC.Detectors.GetDetectorByNumber(1)
        self.QLENGTH = QLENGTH
        self.RHOR = RHOR
        self.redTime = 3.0
        self.RedTimer = 0.0
        self.greenTime = 2.0
        self.GreenTimer = 0.0


    def get_state(self):
        '''
        get current state of the signal light of the ramp meter
        '''
        return self.SG.AttValue('STATE')

    def get_detector(self):
        '''
        get the detector status of the ramp meter
        '''
        return self.Det.AttValue('PRESENCE')

    def red_increment(self, sim_step):
        '''
        increment the red timer
        sim_step: (double) time period of simulation step in seconds
        '''
        self.RedTimer += sim_step

    def change2green(self):
        '''
        change the light to green
        '''
        SG.SetAttValue('STATE', 3)
        RedTimer = 0

    def green_increment(self, sim_step):
        '''
        increment the green time
        sim_step: (double) time period of simulation step in seconds
        '''
        self.GreenTimer += sim_step

    def change2red(self):
        SG.SetAttValue('STATE', 1)
        GreenTimer = 0

    def update_rate(self, rate):
        '''
        update the signal period according to ramp metering rate
        '''
        self.redTime = 3600 / rate - greenTime

    def meter_step(self):
        '''
        update the meter states for one simulation step
        '''
        if self.get_state() == 1:
            self.red_increment()
            if self.get_detector() == 1 and self.RedTimer >= redTime:
                self.change2green()
        else:
            self.green_increment()
            if GreenTimer >= greenTime:
                self.change2red

