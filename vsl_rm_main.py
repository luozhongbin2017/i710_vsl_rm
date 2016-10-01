# attempt to use VISSIM COM interface through Pytyhon win32com
# ref : Mark Hammond, Andy Robinson (2000, Jan). Python Programming on Win32, [Online]. Available: http://www.icodeguru.com/WebServer/Python-Programming-on-Win32/
# ref : VISSIM 5.30-04 - COM Interface Manual, PTV group, Karlsrueh, Germany, 2011.

import ctypes.wintypes
#import email_interface
import inspect
import os
import pywintypes
import win32api
import win32com.client as com
#import win_interface

def getProgramName():
    # ref : In Python, how do I get the path and name of the file that is currently executing?, 2008, [Online] Available at : http://stackoverflow.com/questions/50499/in-python-how-do-i-get-the-path-and-name-of-the-file-that-is-currently-executin
    PROGRAM_FULL_NAME = inspect.getfile(inspect.currentframe())
    PROGRAM_NAME = os.path.split(PROGRAM_FULL_NAME)[-1]
    return PROGRAM_NAME

def mkdir(path):
    ''' recursively create given path '''
    upper, name = os.path.split(path)
    if not os.path.exists(upper):
        mkdir(upper)
    # if path already exists, do not create. Otherwise, error will occur
    if not os.path.exists(path):
        os.mkdir(path)  


def calcRandomSeeds(nMonteCarlo):
    # list comprehension
    # same as {3i + 1 | 0 <= i < nMonteCarlo}
    return [3*i + 1 for i in range(nMonteCarlo)]

def writeHeader(fileName, title, Nsec, firstColumnHeader = "t/s", columnHeader = "Sec"):
    '''write header line of a data file
    Parameters
    ----------
    fileName : string
        file name
    title : string
        first line to be written
    Nsec : int
        number of columns
    firstColumnHeader : string
        header of the first column
    columnHeader : string
        header of each column. Repeated Nsec times with headers
        example : Nsec = 2, columnHeader = "Sec"
        -> "Sec1\tSec2\t"
    '''
    speedFile = open(fileName, 'w')
    # first line
    speedFile.write(title)
    speedFile.write('\n')
    # column headers
    #   first column header
    speedFile.write(firstColumnHeader)
    speedFile.write('\t')

    columnHeaderTemplate = columnHeader+"%d\t"
    # each column
    for i in xrange(Nsec):
        speedFile.write(columnHeaderTemplate % (i + 1))
    speedFile.write('\n')
    speedFile.close()

def runSimulation(simulationTime_sec, idxScenario, idxController, idxLaneClosure, randonSeed, folderDir, networkDir):
    '''run PTV VISSIM simulation using given arguments
    Parameters
    ----------
        simulationTime_sec : int
            Simulation time (s)
        startTime_sec : int
            Time accident starts (s)
        endTime_sec : int
            Time accident ends (s)
        vehicleInput_veh_hr : int
            vehicle input (veh/hr)
        idxScenario : int
            Index of simulation scenarios
        idxController : int
            Controller Index
        idxLaneClosure : int
            Whether Lane Change control is added
        randomSeed : int
            seed for random numbers
        folderDir : string, path
            location of data files
        networkDir : string, path
            location of network file
    '''

    # Link Groups
    link_groups = [
                {'MAINLINE': (84,),         'ONRAMP': (),       'OFFRAMP': (),      'DC': 1, 'VSL': (1, 2, 3, 4)},
                {'MAINLINE': (83,),         'ONRAMP': (),       'OFFRAMP': (),      'DC': 2, 'VSL': (5, 6, 7, 8)},
                {'MAINLINE': (81,),         'ONRAMP': (82,),    'OFFRAMP': (80,),   'DC': 3, 'VSL': (9, 10, 11, 12, 13, 69) },
                {'MAINLINE': (79,),         'ONRAMP': (),       'OFFRAMP': (),      'DC': 6, 'VSL': (14, 15, 16, 17)},
                {'MAINLINE': (77, 76),      'ONRAMP': (78,),    'OFFRAMP': (),      'DC': 7, 'VSL': (18, 19, 20, 21, 70)},
                {'MAINLINE': (75, 74),      'ONRAMP': (),       'OFFRAMP': (73,),   'DC': 9, 'VSL': (22, 23, 24, 25)},
                {'MAINLINE': (72, 70),      'ONRAMP': (71,),    'OFFRAMP': (),      'DC': 11, 'VSL': (26, 27, 28, 29, 72)},
                {'MAINLINE': (69,),         'ONRAMP': (),       'OFFRAMP': (),      'DC': 13, 'VSL': (30, 31, 32, 33)},
                {'MAINLINE': (68, 66,),     'ONRAMP': (),       'OFFRAMP': (67,),   'DC': 14, 'VSL': (34, 35, 36, 37)},
                {'MAINLINE': (65,),         'ONRAMP': (),       'OFFRAMP': (),      'DC': 16, 'VSL': (38, 39, 40)},
                {'MAINLINE': (63, 62, 60,), 'ONRAMP': (64,),    'OFFRAMP': (61,),   'DC': 17, 'VSL': (41, 42, 43, 44, 75)},
                {'MAINLINE': (59,),         'ONRAMP': (),       'OFFRAMP': (),      'DC': 20, 'VSL': (45, 46, 47)},
                {'MAINLINE': (58,),         'ONRAMP': (),       'OFFRAMP': (),      'DC': 21, 'VSL': (48, 49, 50)},
                {'MAINLINE': (57, 55, 51),  'ONRAMP': (54,),    'OFFRAMP': (56, 53), 'DC': 22, 'VSL': (51, 52, 53, 54, 77)},
                {'MAINLINE': (50, 19, 18),  'ONRAMP': (49,),    'OFFRAMP': (),      'DC': 26, 'VSL': (55, 56, 57, 79)},
                {'MAINLINE': (17, 339),     'ONRAMP': (),       'OFFRAMP': (326,),  'DC': 28, 'VSL': (58, 59, 60)},
                {'MAINLINE': (338,),        'ONRRAMP': (),      'OFFRAMP': (),      'DC': 30, 'VSL': (61, 62, 63)},
                {'MAINLINE': (337, 310, 311, 309, 307, 306, 304), 
                                            'ONRAMP': (336, 308),'OFFRAMP': (312, 305), 'DC': 31, 'VSL': (65, 66, 67, 81)}
    ]

    #  Ramps
    ramps = {
                82: {'LINK_GROUP': 2, 'TYPE': 'ON', 'DC': 4, 'SC': 1, 'QC': 1, 'QLENGTH': 100, 'RHOR': 90},
                80: {'LINK_GROUP': 2, 'TYPE': 'OFF', 'DC': 5},
                78: {'LINK_GROUP': 4, 'TYPE': 'ON', 'DC': 8, 'SC': 2, 'QC': 2, 'QLENGTH': 100, 'RHOR': 90},
                73: {'LINK_GROUP': 5, 'TYPE': 'OFF', 'DC': 10},
                71: {'LINK_GROUP': 6, 'TYPE': 'ON', 'DC': 12, 'SC': 3, 'QC': 3, 'QLENGTH': 100, 'RHOR': 90},
                67: {'LINK_GROUP': 8, 'TYPE': 'OFF', 'DC': 15},
                64: {'LINK_GROUP': 10, 'TYPE': 'ON', 'DC': 18, 'SC': 4, 'QC': 4, 'QLENGTH': 100, 'RHOR': 90},
                61: {'LINK_GROUP': 10, 'TYPE': 'ON', 'DC': 19, 'SC': 5, 'QC': 5, 'QLENGTH': 100, 'RHOR': 90},
                56: {'LINK_GROUP': 13, 'TYPE': 'OFF', 'DC': 24},
                54: {'LINK_GROUP': 13, 'TYPE': 'ON', 'DC': 23, 'SC': 6, 'QC': 6, 'QLENGTH': 100, 'RHOR': 90},
                53: {'LINK_GROUP': 13, 'TYPE': 'OFF', 'DC': 25},
                49: {'LINK_GROUP': 14, 'TYPE': 'ON', 'DC': 27, 'SC': 7, 'QC': 7, 'QLENGTH': 100, 'RHOR': 90},
                326:{'LINK_GROUP': 15, 'TYPE': 'OFF', 'DC': 29},
                336:{'LINK_GROUP': 17, 'TYPE': 'ON', 'DC': 32, 'SC': 8, 'QC': 8, 'QLENGTH': 100, 'RHOR': 90},
                312:{'LINK_GROUP': 17, 'TYPE': 'OFF', 'DC': 34},
                #308:{'LINK_GROUP': 17, 'TYPE': 'ON', 'DC': 33},
                305:{'LINK_GROUP': 17, 'TYPE': 'OFF', 'DC': 35},
    }


    '''setting of scenarios'''
    scenarios = [{'group': 17, 'link': 306, 'lane': 2, 'coordinate': 5, startTime_sec: 1500, endTime_sec: 2100},
                 {'group': 17, 'link': 306, 'lane': 2, 'coordinate': 5, startTime_sec: 1500, endTime_sec: 3300},
                 {'group': 17, 'link': 306, 'lane': 2, 'coordinate': 5, startTime_sec: 1500, endTime_sec: 4800}]


    ''' Demand settings'''
    #demands = [4500, 280, 180, 400, 650, 10, 50, 180, 100, 50, 90, 280, 550] # daily average vehicle demands
    demands = [4500, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    for i in xrange(len(demands)):
        demands[i] *= demandRatio


    Nsec = len(link_groups)   # number of sections
    Ninput = len(demands)   # number of vehicle input points
    simResolution = 5   # No. of simulation steps per second
    stepTime_sec = 1.0 / simResolution  # length of single simulation step in seconds
    Tdata_sec = 5.0     # Data sampling period
    Tctrl_sec = 30.0 # Control time interval


    car_class = 10
    truck_class = 20
    # vf = {car_class: 65, truck_class: 65} #default speed limit of cars and trucks
    vf = 65 # The free flow speed of highway in mi/h

    fts2mph = 0.681818 # 1 ft/s = 0.681818 mph

    # speed[]: average speed of each section
    # density[]: vehicle density of each section
    # flowSection[]: flow rate of each section
    # flowLane[]: flow rate of each lane at bottleneck section
    # vsl[]: vsl command of each section at current control period, a list of dicts
    # vslOld[]: vsl command of each section at previous control period
    speed = [0.0] * Nsec
    density = [0.0] * Nsec
    #densityOld = [0.0] * Nsec
    flowSection = [0.0] * Nsec
    vsl = [vf] * Nsec
    vslOld = [vf] * Nsec


    '''Define log file names'''
    speedFileName = os.path.join(folderDir, "speedLog_%d.txt"%randomSeed)
    densityFileName = os.path.join(folderDir, "densityLog_%d.txt"%randomSeed)
    flowSectionFileName = os.path.join(folderDir, "flowSectionLog_%d.txt"%randomSeed)
    vslFileName = os.path.join(folderDir, "vslLog_%d.txt"%randomSeed)


    '''write file headers'''
    writeHeader(speedFileName, "Average Speed of Each Section", Nsec)
    writeHeader(densityFileName, "Density of Each Section", Nsec)
    writeHeader(flowSectionFileName, "Flow Rate of Each Section", Nsec)
    writeHeader(vslFileName, "VSL Command of Each Section", Nsec)

    
    ProgID = "VISSIM.Vissim"
    
    '''file paths'''
    networkFileName = "I710.inp"
    layoutFileName = "I710.ini"
    #path = networkDir
    #arrivedVehFileName = os.path.join(path, "ArrivedVehs.txt")
    
    ''' Start VISSIM simulation '''
    ''' COM lines'''  

    try:
        print "Client: Creating a Vissim instance"
        vs = com.Dispatch(ProgID)
        
        print "Client: Creating a Vissim Simulation instance"
        sim = vs.Simulation
    
        print "Client: read network and layout"
        vs.LoadNet(os.path.join(networkDir, networkFileName), 0)
        vs.LoadLayout(os.path.join(networkDir, layoutFileName))
    
        ''' initialize simulations '''
        ''' setting random seed, simulation time, simulation resolution, simulation speed '''
        print "Client: Setting simulations"
        sim.Resolution = simResolution
        sim.Speed = 0
        sim.RandomSeed = randomSeed

        # Reference to road network
        net = vs.Net

        ''' Set vehicle input '''
        vehInputs = net.VehicleInputs

        for num in xrange(Ninput):
            vehInput = vehInputs.GetVehicleInputByNumber(num+1)
            vehInput.SetAttValue('VOLUME', demands[num])
            vehInput.SetAttValue('TRAFFICCOMPOSITION', demandComposition)                

        ''' Set default speed limit '''
        VSLs = net.DesiredSpeedDecisions
        numberSpeedLimit = VSLs.count
        for VSL in VSLs :
            VSL.SetAttValue1("DESIREDSPEED", car_class, vf)
            VSL.SetAttValue1("DESIREDSPEED", truck_class, vf)

        ''' Get Data collection and link objects '''
        dataCollections = net.DataCollections

        # Make sure that all lanes are open        
        links = net.Links
        link_Nlane = {}
        link_length = {}
        for link in links:
            ID = link.ID
            link_Nlane[ID] = link.AttValue('NUMLANES')
            link_length[ID] = link.AttValue('LENGTH')
            # Open all lanes
            for lane in range(link_Nlane[ID]):
                link.SetAttValue2('LANECLOSED', lane+1, car_class, False)
                link.SetAttValue2('LANECLOSED', lane+1, truck_class, False)

        busNo = 2 # No. of buses used to block the lane
        busArray = [0] * busNo

        vehs = net.Vehicles


        # start simulation
        print "Client: Starting simulation"
        print "Vehicle Input Ratio:", demandRatio
        print "Scenario:", idxScenario
        print "Controller:", idxController
        print "Lane Closure:", idxLaneClosure, "section"
        print "Random Seed:", randomSeed

        





    except pywintypes.com_error, err:
        print "err=", err
        var0 = err[0]
        print "%d == 0x%x" % (var0, ctypes.wintypes.UINT(var0).value), win32api.FormatMessage(var0)
        # https://mail.python.org/pipermail/python-win32/2008-August/008041.html
        print 
        # DISP_E_EXCEPTION == 0x80020009
        
        var1 = err[2][-2]
        print "%d == 0x%x" % (var1, ctypes.wintypes.UINT(var1).value)
        # DISP_E_EXCEPTION == 0x80020009

        var2 = err[2][-1]
        print "%d == 0x%x" % (var2, ctypes.wintypes.UINT(var2).value), win32api.FormatMessage(var2)
        # real error code
        # E_FAIL == 0x80004005
        # ref https://mail.python.org/pipermail/python-win32/2008-August/008040.html
        #     http://www.vistax64.com/vista-installation-setup/33219-regsvr32-error-0x80004005.html
        #     http://sharepoint.stackexchange.com/questions/42838/exception-from-hresult-0x80020009-disp-e-exception
        
        # ref VISSIM_COM p.236 Error Handling
        #     VISSIM_COM p.258 Error Messages list of error messages by VISSIM server
        #     VISSIM_COM p.262 "The specified configuration is not defined within VISSIM."
        '''
        Error
        The specified configuration is not defined within VISSIM.
        
        Description
        Some methods for evaulations results need a previously configuration for data collection. 
        The error occurs when requesting results that have not been previously configured.
        For example, using the GetSegmentResult() method of the ILink interface to request
        density results can end up with this error if the density has not been requested within the configuration
        ''' 
        
        #email_interface.mail(DEST, getProgramName(), "exception : %s" % (err)) 
    
    except Exception, err:
        print "err=", err
        #email_interface.mail(DEST, getProgramName(), "exception : %s" % (err)) 
# end runSimulation()




    
