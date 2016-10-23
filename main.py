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
import shutil
import RampMeter as RM
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
    LogFile = open(fileName, 'w')
    # first line
    LogFile.write(title)
    LogFile.write('\n')
    # column headers
    #   first column header
    LogFile.write(firstColumnHeader)
    LogFile.write('\t')

    columnHeaderTemplate = columnHeader+"%d\t"
    # each column
    for i in xrange(Nsec):
        LogFile.write(columnHeaderTemplate % (i + 1))
    LogFile.write('\n')
    LogFile.close()

def writeRampHeader(fileName, title, IDList, firstColumnHeader = "t/s", columnHeader = "Ramp"):
    '''write header line of a data file
    Parameters
    ----------
    fileName : string
        file name
    title : string
        first line to be written
    IDList : list[int]
        List of ID of the ramps
    firstColumnHeader : string
        header of the first column
    columnHeader : string
        header of each column. Repeated Nsec times with headers
        example : Nsec = 2, columnHeader = "Sec"
        -> "Sec1\tSec2\t"
    '''
    LogFile = open(fileName, 'w')
    # first line
    LogFile.write(title)
    LogFile.write('\n')
    # column headers
    #   first column header
    LogFile.write(firstColumnHeader)
    LogFile.write('\t')

    columnHeaderTemplate = columnHeader+"%d\t"
    # each column
    for ID in IDList:
        LogFile.write(columnHeaderTemplate % ID)
    LogFile.write('\n')
    LogFile.close()

def LC_Control(scenario, links_obj, LC_distance = 2000):
    '''Apply Lane Change control for specific scenario
    Parameters
    -------------
        scenario: dict
            the scenario dictionary
        links_obj: links object of VISSIM
            the links object
        link_groups: list
            the list of link groups
        ramp_links: list
            the list of links with off ramp
        LC_distance: int 
        the distance of LC controlled section (ft), default 2000

    '''
    car_class = 10
    truck_class = 20
    length = 0

    link_list = [84, 10103, 83, 10102, 81, 10099, 79, 10098, 77, 10096, 76, 10095, 75, 10094, 74, 10092, 72, 10091, 70, 10089, 69, 10087, 68, 10088, 66, 10085, 65, 10084, 63, 10082, 62, 10081, 60, 10079, 59, 10078, 58, 10077, 57, 10076, 55, 10074, 51, 10070, 50, 10030, 19, 10023, 18, 10022, 17, 10021, 339, 10341, 338, 10339, 337, 10337, 310,10311, 311, 10313, 309, 10310, 307, 10308, 306, 10306, 304]
    link_ID = scenario['link']
    link_idx = link_list.index(link_ID)
    lane = scenario['lane']
    link_obj = links_obj.GetLinkByNumber(link_ID)
    Nlane = link_obj.AttValue('NUMLANES')
    while length < LC_distance:            
        link_obj.SetAttValue2('LANECLOSED', lane, car_class, True)
        link_obj.SetAttValue2('LANECLOSED', lane, truck_class, True)
        length += link_obj.AttValue('LENGTH')
        link_idx -= 1
        link_ID = link_list[link_idx]
        link_obj = links_obj.GetLinkByNumber(link_ID)
        lane = link_obj.AttValue('NUMLANES') - (Nlane - lane)
        Nlane = link_obj.AttValue('NUMLANES')
    else:
        print LC_distance, 'ft of lane closed'


def alineaQ(rmRate, density, q, demand, RHOR, QLENGTH):
    rmMIN = 100
    rmMAX = 1200
    alpha = 1
    beta = 1
    rm_den = rmRate - alpha * (density - RHOR)
    rm_Q = demand + beta * (q - QLENGTH)
    return max(rmMIN, min(rmMAX, max(rm_den, rm_Q)))

def vsl_FeedbackLinearization(density, vsl, section_length, rmRate, rampFlow, link_groups):
    '''
    density: (list[Nsec]) list of densities in all sections
    vsl: (list[Nsec]) current VSL command in each sections
    section_length: (list[Nsec]) length of each section
    
    '''
    rampDemmand = {64: 200, 61:500, 54:200, 49:300} # use to estimate the on-ramp flow
    startSection = 9    #startSection: (int) the first section controlled with VSL
    endSection = 15    #endSection: (int) the last section upstream the discharging section. All sections downstream of endSection are discharging sections
    Nsec_controlled = endSection - startSection + 2 # Nsec_controlled is the number of sections under control, including the discharging section! The discharging section is considered as one single section

    car_class = 10
    truck_class = 20

    ve = 40
    vslMIN = 10
    vslMAX = 65

    # Set desired equilibrium point
    wb = 40
    rho_e = [120.0] * Nsec_controlled
    rho_e[0] = 430.0
    v_e = [ve] * (Nsec_controlled - 1)
    v_e[0] = 15

    # The length of 
    rho = density[startSection:(endSection + 2)]
    L = section_length[startSection:(endSection + 2)]

    #compute the average density and total length in the discharging section
    rho[-1] = 0.0
    L[-1] = 0.0
    for idx in range((endSection+1), len(density)):
        rho[-1] += density[idx] * section_length[idx]
        L[-1] += section_length[idx]
    rho[-1] = rho[-1] / L[-1]

    # make sure that the control command is not divided by zero
    for idx in range(len(rho)):
        rho[idx] = max(rho[idx], 10)


    # compute the on-ramp flow and off-ramp flow of each section
    onFlow = [0.0] * Nsec_controlled
    offFlow = [0.0] * Nsec_controlled
    for iSec in range(Nsec_controlled):
        for jRamp in link_groups[iSec + startSection]['ONRAMP']:
            onFlow[iSec] += min(rmRate[jRamp], rampDemmand[jRamp])
        for jRamp in link_groups[iSec + startSection]['OFFRAMP']:
            offFlow[iSec] += rampFlow[jRamp]

    # Get the error state
    e = [0.0] * Nsec_controlled
    for i in range(len(e)):
        e[i] = rho[i] - rho_e[i]

    # Control parameter Lambda
    Lambda = [20] * (Nsec_controlled - 1)

    # vsl command
    v = [0.0] * (Nsec_controlled - 1)

    for i in xrange(len(v) - 1):
        v[i] = (-Lambda[i] * L[i] * e[i+1] + ve * rho_e[-1] - sum(onFlow[(i+1):]) + sum(offFlow[(i+1):])) / rho[i]
        
    if e[-1] <= 0:
        v[-1] = (-Lambda[-1] * L[-1] * e[-1] + ve * rho[-1] - onFlow[-1] + offFlow[-1]) / rho[-1]
    else:
        v[-1] = (-Lambda[-1] * L[-1] * e[-1] + ve * rho_e[-1] - wb * e[-1]   - onFlow[-1] + offFlow[-1]) / rho[-1]

    for i in xrange(Nsec_controlled - 1):
        v[i] = round(v[i] * 0.2) * 5
        if v[i] <= vsl[i + startSection]:
            v[i] = max(v[i], vsl[i + startSection] - 10, vslMIN)
        else:
            v[i] = min(v[i], vslMAX)
        vsl[i+startSection] = v[i]

    return vsl



    










    



def runSimulation(simulationTime_sec, idxScenario, idxController, idxLaneClosure, randomSeed, folderDir, networkDir, demandRatio, demandComposition):
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
                {'MAINLINE': (338,),        'ONRAMP': (),      'OFFRAMP': (),      'DC': 30, 'VSL': (61, 62, 63)},
                {'MAINLINE': (337, 310, 311, 309, 307, 306, 304), 
                                            'ONRAMP': (336, 308),'OFFRAMP': (312, 305), 'DC': 31, 'VSL': (65, 66, 67, 81)}
    ]

    #  Ramps
    ramps = {
                82: {'LINK_GROUP': 2, 'TYPE': 'ON', 'INPUT': 2, 'DC': 4, 'SC': 1, 'QC': 1, 'QLENGTH': 100, 'RHOR': 90},
                80: {'LINK_GROUP': 2, 'TYPE': 'OFF', 'DC': 5},
                78: {'LINK_GROUP': 4, 'TYPE': 'ON', 'INPUT': 3, 'DC': 8, 'SC': 2, 'QC': 2, 'QLENGTH': 100, 'RHOR': 90},
                73: {'LINK_GROUP': 5, 'TYPE': 'OFF', 'DC': 10},
                71: {'LINK_GROUP': 6, 'TYPE': 'ON', 'INPUT': 4, 'DC': 12, 'SC': 3, 'QC': 3, 'QLENGTH': 100, 'RHOR': 90},
                67: {'LINK_GROUP': 8, 'TYPE': 'OFF', 'DC': 15},
                64: {'LINK_GROUP': 10, 'TYPE': 'ON', 'INPUT': 5, 'DC': 18, 'SC': 4, 'QC': 4, 'QLENGTH': 100, 'RHOR': 90},
                61: {'LINK_GROUP': 10, 'TYPE': 'ON', 'INPUT': 6, 'DC': 19, 'SC': 5, 'QC': 5, 'QLENGTH': 100, 'RHOR': 90},
                56: {'LINK_GROUP': 13, 'TYPE': 'OFF', 'DC': 24},
                54: {'LINK_GROUP': 13, 'TYPE': 'ON', 'INPUT': 7, 'DC': 23, 'SC': 6, 'QC': 6, 'QLENGTH': 100, 'RHOR': 90},
                53: {'LINK_GROUP': 13, 'TYPE': 'OFF', 'DC': 25},
                49: {'LINK_GROUP': 14, 'TYPE': 'ON', 'INPUT': 8, 'DC': 27, 'SC': 7, 'QC': 7, 'QLENGTH': 100, 'RHOR': 90},
                326:{'LINK_GROUP': 15, 'TYPE': 'OFF', 'DC': 29},
                336:{'LINK_GROUP': 17, 'TYPE': 'ON', 'INPUT': 9, 'DC': 32, 'SC': 8, 'QC': 8, 'QLENGTH': 100, 'RHOR': 90},
                312:{'LINK_GROUP': 17, 'TYPE': 'OFF', 'DC': 34},
                #308:{'LINK_GROUP': 17, 'TYPE': 'ON', 'DC': 33},
                305:{'LINK_GROUP': 17, 'TYPE': 'OFF', 'DC': 35},
    }


    '''setting of scenarios'''
    scenarios = [{'group': 17, 'link': 306, 'lane': 2, 'coordinate': 5, 'startTime_sec': 600, 'endTime_sec': 1200},
                 {'group': 17, 'link': 306, 'lane': 2, 'coordinate': 5, 'startTime_sec': 1500, 'endTime_sec': 3300},
                 {'group': 17, 'link': 306, 'lane': 2, 'coordinate': 5, 'startTime_sec': 1500, 'endTime_sec': 4800}]


    ''' Demand settings'''
    #demands = [4500, 280, 180, 400, 650, 10, 50, 180, 100, 50, 90, 280, 550] # daily average vehicle demands
    demands = [4500, 0, 0, 0,200, 500, 200, 300, 0, 0]
    #demands = [4500, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    for i in xrange(len(demands)):
        demands[i] *= demandRatio


    Nsec = len(link_groups)   # number of sections
    Ninput = len(demands)   # number of vehicle input points
    Nramp = len(ramps)  # number of ramps
    simResolution = 5   # No. of simulation steps per second
    stepTime_sec = 1.0 / simResolution  # length of single simulation step in seconds
    Tdata_sec = 5.0     # Data sampling period
    Tctrl_sec = 30.0 # Control time interval

    #   Incident time period
    startTime_sec = scenarios[idxScenario]['startTime_sec']
    endTime_sec = scenarios[idxScenario]['endTime_sec']


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
    flowSection = [0.0] * Nsec
    vsl = [vf] * Nsec
    vslOld = [vf] * Nsec
    rampFlow = {}
    rampQue = {}
    rmRate = {}
    for key in ramps:
        rampFlow[key] = 0.0
        rampQue[key] = 0.0
        rmRate[key] = 0.0
    


    '''Define log file names'''
    speedFileName = os.path.join(folderDir, "speedLog_%d.txt"%randomSeed)
    densityFileName = os.path.join(folderDir, "densityLog_%d.txt"%randomSeed)
    flowSectionFileName = os.path.join(folderDir, "flowSectionLog_%d.txt"%randomSeed)
    vslFileName = os.path.join(folderDir, "vslLog_%d.txt"%randomSeed)
    rampFlowFileName = os.path.join(folderDir, "rampFlowLog_%d.txt"%randomSeed)
    rampQueFileName = os.path.join(folderDir, "rampQueLog_%d.txt"%randomSeed)
    rmRateFileName = os.path.join(folderDir, "rmRateLog_%d.txt"%randomSeed)


    '''write file headers'''
    writeHeader(speedFileName, "Average Speed of Each Section", Nsec)
    writeHeader(densityFileName, "Density of Each Section", Nsec)
    writeHeader(flowSectionFileName, "Flow Rate of Each Section", Nsec)
    writeHeader(vslFileName, "VSL Command of Each Section", Nsec)
    writeRampHeader(rampFlowFileName, "Flow rate on each ramp", ramps.keys())
    writeRampHeader(rampQueFileName, "Queue Length on each ramp", ramps.keys())
    writeRampHeader(rmRateFileName, "Ramp Metering Rate of each ramp", ramps.keys())


    
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
        links = net.Links

        # Make sure that all lanes are open      
        link_Nlane = {} # length of each link
        link_length = {}    # number of lanes in each link
        for link in links:
            ID = link.ID
            link_Nlane[ID] = link.AttValue('NUMLANES')
            link_length[ID] = link.AttValue('LENGTH')
            # Open all lanes
            for lane in range(link_Nlane[ID]):
                link.SetAttValue2('LANECLOSED', lane+1, car_class, False)
                link.SetAttValue2('LANECLOSED', lane+1, truck_class, False)
        section_length = [0.0] * Nsec # length of each section
        for iSec in xrange(len(link_groups)):
            for linkID in link_groups[iSec]['MAINLINE']:
                link = links.GetLinkByNumber(linkID)
                section_length[iSec] += link.AttValue('LENGTH') / 5280



        busNo = 2 # No. of buses used to block the lane
        busArray = [0] * busNo

        vehs = net.Vehicles
        signalControllers = net.SignalControllers #VISSIM SignalControllers object
        queueCounters = net.QueueCounters #VISSIM queue counter object

        # Construct the dictionary of ramp meters
        meters_obj = {}
        for key in ramps.keys():
            ramp = ramps[key]
            if ramp['TYPE'] == 'ON':
                meters_obj[key] = RM.RampMeter(ramp['LINK_GROUP'], vehInputs.GetVehicleInputByNumber(ramp['INPUT']), dataCollections.GetDataCollectionByNumber(ramp['DC']), signalControllers.GetSignalControllerByNumber(ramp['SC']), queueCounters.GetQueueCounterByNumber(ramp['QC']), ramp['QLENGTH'], ramp['RHOR'])
                rmRate[key] = max(meters_obj[key].INPUT.AttValue('VOLUME'), 100)
                meters_obj[key].update_rate(rmRate[key])





        # start simulation
        print "Client: Starting simulation"
        print "Vehicle Input Ratio:", demandRatio
        print "Scenario:", idxScenario
        print "Controller:", idxController
        print "Lane Closure:", idxLaneClosure, "section"
        print "Random Seed:", randomSeed

        currentTime = 0
        while currentTime <= simulationTime_sec:
            # run sigle step of simulation
            sim.RunSingleStep()
            currentTime = sim.AttValue('ELAPSEDTIME')
            for key in meters_obj:
                meters_obj[key].meter_step(stepTime_sec)
            #currentTime += stepTime_sec
            print currentTime
            if 0 == currentTime % Tdata_sec:
                # Get density, flow, speed of each section
                for iSection in xrange(Nsec):
                    # Flow volume in one sampling period
                    #print link_groups[iSection]['DC']
                    dataCollection = dataCollections.GetDataCollectionByNumber(link_groups[iSection]['DC'])
                    flowSection[iSection] = dataCollection.GetResult('NVEHICLES', 'sum', 0) 

                    # Density and Speed in each section
                    Nveh = 0
                    dist = 0
                    denomenator = 0
                    for linkID in link_groups[iSection]['MAINLINE']:
                        link = links.GetLinkByNumber(linkID)
                        den = link.GetSegmentResult("DENSITY", 0, 0, 1, 0)
                        v = link.GetSegmentResult("SPEED",   0, 0, 1, 0)
                        Nveh += den * link_length[linkID]
                        dist += v * den * link_length[linkID] #total distance tranveled by all vehicles in one link

                        denomenator += link_length[linkID]
                    density[iSection] = Nveh / denomenator
                    if 0 == Nveh:
                        speed[iSection] = 0
                    else:
                        speed[iSection] = dist / Nveh


                        
                ''' write log files : flowSection, speed, density '''
                #   open files and indicate time
                flowSectionFile = open(flowSectionFileName, "a")
                flowSectionFile.write(str(currentTime) + '\t')
                
                speedFile = open(speedFileName, "a")
                speedFile.write(str(currentTime) + '\t')

                densityFile = open(densityFileName, "a")
                densityFile.write(str(currentTime) + '\t')
                    
                #   write values. Nsec loop
                for fl, sp, de in zip(flowSection, speed, density):
                    flowSectionFile.write(str(fl) + '\t')
                    speedFile.write(str(sp) + '\t')
                    densityFile.write(str(de) + '\t')
                # end for fl, sp, de in zip(flowSection, speed, density)

                # mark end of line and close file
                flowSectionFile.write('\n'); flowSectionFile.close()
                speedFile.write('\n'); speedFile.close()
                densityFile.write('\n'); densityFile.close()
                # end write log files : flowSection, speed, density

                # Get flow on each ramp
                rampFlowFile = open(rampFlowFileName, 'a')
                rampFlowFile.write(str(currentTime) + '\t')
                for key in ramps.keys():
                    dataCollection = dataCollections.GetDataCollectionByNumber(ramps[key]['DC'])
                    flow = dataCollection.GetResult('NVEHICLES', 'sum', 0) 
                    rampFlowFile.write(str(flow) + '\t')
                    rampFlow[key] += flow
                rampFlowFile.write('\n'); rampFlowFile.close()





            if currentTime == startTime_sec:
                # set incident scenario
                for i in xrange(busNo):
                    busArray[i] = vehs.AddVehicleAtLinkCoordinate(300, 0, scenarios[idxScenario]['link'], scenarios[idxScenario]['lane'], scenarios[idxScenario]['coordinate'] + 40 * i, 0)
                # Apply Lane change control
                LC_Control(scenarios[idxScenario], links, 2000 * idxLaneClosure)

            if currentTime == endTime_sec:
                # Remove the incident
                for veh in busArray:
                    vehs.RemoveVehicle(veh.ID)

                # open all closed lanes
                for link in links:
                    ID = link.ID
                    for lane in xrange(link_Nlane[ID]):
                        link.SetAttValue2('LANECLOSED', lane+1, car_class, False)
                        link.SetAttValue2('LANECLOSED', lane+1, truck_class, False)

            #Compute the VSL command
            if 0 == currentTime % Tctrl_sec:
                for key in meters_obj.keys():
                    rampQue[key] = meters_obj[key].QC.GetResult(currentTime, 'MEAN')

                rampQueFile = open(rampQueFileName, 'a')
                rampQueFile.write(str(currentTime) + '\t')
                for key in ramps.keys():
                    rampQueFile.write(str(rampQue[key]) + '\t')
                rampQueFile.write('\n')
                rampQueFile.close()

                for key in meters_obj.keys():
                    ramp = meters_obj[key]
                    rmRate[key] = alineaQ(rmRate[key], density[ramp.LINK_GROUP], rampQue[key], meters_obj[key].INPUT.AttValue('VOLUME'), meters_obj[key].RHOR, meters_obj[key].QLENGTH)
                    ramp.update_rate(rmRate[key])

                rmRateFile = open(rmRateFileName, 'a')
                rmRateFile.write(str(currentTime) + '\t')
                for key in ramps.keys():
                    rmRateFile.write(str(rmRate[key]) + '\t')
                rmRateFile.write('\n')
                rmRateFile.close()

                
                if (startTime_sec <= currentTime < endTime_sec):
                    #pass
                    vsl = vsl_FeedbackLinearization(density, vsl, section_length, rmRate, rampFlow, link_groups)
                else:
                    for iVSL in xrange(Nsec):
                        vsl[iVSL] = vf

                #Update VSL command in VISSIM
                for iSec in xrange(Nsec):
                    for vslID in link_groups[iSec]["VSL"]:
                        VSL = VSLs.GetDesiredSpeedDecisionByNumber(vslID)
                        VSL.SetAttValue1('DESIREDSPEED', car_class, vsl[iSec])
                        VSL.SetAttValue1('DESIREDSPEED', truck_class, vsl[iSec])

                # write the VSL log file
                vslFile = open(vslFileName, 'a')
                vslFile.write(str(currentTime) + '\t')
                for iVSL in xrange(len(vsl)):
                    vslFile.write(str(vsl[iVSL]) + '\t')
                vslFile.write('\n')
                vslFile.close()

                for key in rampFlow:
                    rampFlow[key] = 0.0





            # # run sigle step of simulation
            # sim.RunSingleStep()
            # currentTime = sim.AttValue('ELAPSEDTIME')
            

        





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

def main():
    scenario = 0
    nMonteCarlo = 3
    inc = [0]
    ctrl = [(0, 2)]

    demandRatio = 1.5
    simulationTime_sec = 4800
    randomSeeds = calcRandomSeeds(nMonteCarlo)

    # Vehicle Composition ID
    # 1: 30% Trucks
    # 2: 15% Trucks
    demandComposition = 2

    networkDir = os.path.abspath(os.curdir)

    for iInc in inc:
        for jController, kLaneClosure in ctrl:
            folderDir = os.path.join(networkDir, 'data', 'scenario%d' %iInc, 'ctrl%d' %jController, 'LC%d' %kLaneClosure)
            mkdir(folderDir)
            for lMC in randomSeeds:
                runSimulation(simulationTime_sec, iInc, jController, kLaneClosure, lMC, folderDir, networkDir, demandRatio, demandComposition)
                shutil.copyfile(os.path.join(networkDir, "I710.fzp"), os.path.join(folderDir, "I710_%d"%lMC))



if __name__ == '__main__':
    main()







    
