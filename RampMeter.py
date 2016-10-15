# This file defines the class of ramp meter.

class RampMeter(object):
    """
    LINK_GROUP: (int) the link group index of the section that the ramp meter is conneted to
    DC: (int) the ID of data collection on the ramp
    SC: (int) the ID of the signal controller associated with the ramp meter
    QC: (int) the ID of the queue counter associated with the ramp meter
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
        self.QLENGTH = QLENGTH
        self.RHOR = RHOR


