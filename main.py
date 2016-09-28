import ctypes.wintypes
import inspect
import os
import pywintypes
import win32api
import win32com.client as com
import fzpEvaluation as fzp
import numpy as np
import shutil

def getProgramName():
    # ref : In Python, how do I get the path and name of the file that is currently executing?, 2008, [Online] Available at : http://stackoverflow.com/questions/50499/in-python-how-do-i-get-the-path-and-name-of-the-file-that-is-currently-executin
    PROGRAM_FULL_NAME = inspect.getfile(inspect.currentframe())
    PROGRAM_NAME = os.path.split(PROGRAM_FULL_NAME)[-1]
    return PROGRAM_NAME

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
    logFile = open(fileName, 'w')
    # first line
    logFile.write(title)
    logFile.write('\n')
    # column headers
    #   first column header
    logFile.write(firstColumnHeader)
    logFile.write('\t')

    columnHeaderTemplate = columnHeader+"%d\t"
    # each column
    for i in xrange(Nsec):
        logFile.write(columnHeaderTemplate % (i + 1))
    
    logFile.write('\n')
    logFile.close()   

def LC_Control(scenario, links_obj, link_groups, ramp_links, LC_distance = 2000):
    '''Apply Lane Change control for specific scenario
    Parameters
    -------------
        scenario: dict
            the scenario dictionary
        links: links object of VISSIM
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

    #link_list = [item for tup in link_groups for item in tup]
    link_list = [84, 10103, 83, 10102, 81, 10099, 79, 10098, 77, 10096, 76, 10095, 75, 10094, 74, 10092, 72, 10091, 70, 10089, 69, 10087, 68, 10088, 66, 10085, 65, 10084, 63, 10082, 62, 10081, 60, 10079, 59, 10078, 58, 10077, 57, 10076, 55, 10074, 51, 10070, 50, 10030, 19, 10023, 18, 10022, 17, 10021, 339, 10341, 338, 10339, 337, 10337, 310,10311, 311, 10313, 309, 10310, 307, 10308, 306, 10306, 304]
    link_ID = scenario['link']
    link_idx = link_list.index(link_ID)
    lane = scenario['lane']
    link_obj = links_obj.GetLinkByNumber(link_ID)
    Nlane = link_obj.AttValue('NUMLANES')
    while length < LC_distance:
        if True: #link_ID not in ramp_links:            
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




def runSimulation(simulationTime_sec, startTime_sec, endTime_sec, demandRatio, demandComposition, idxScenario, idxController, idxLaneClosure, randomSeed, folderDir, networkDir):
    '''run PTV VISSIM simulation using given arguments
    Parameters
    ----------
        simulationTime_sec : int
            Simulation time (s)
        startTime_sec : int
            Time accident starts (s)
        endTime_sec : int
            Time accident ends (s)
        vehInputRatio : double
            Ratio to the daily average vehcle input
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


    # Definition of link/ data collection/ VSL groups
    link_groups = [(84,),
                   (83,),
                   (81,),
                   (79, 77),
                   (76,),
                   (75, 74),
                   (72, 70),
                   (69,),
                   (68, 66,),
                   (65,),
                   (63, 62, 60,),
                   (59,),
                   (58,),
                   (57, 55, 51),
                   (50, 19, 18),
                   (17,),
                   (339, 338, 337,),
                   (310, 311, 309, 307, 306, 304,),
                   ]

    DC_groups = [(1, 2, 3, 4),
                 (5, 6, 7, 8),
                 (9, 10, 11, 12, 13,),
                 (14, 15, 16, 17, 18,),
                 (19, 20, 21, 22),
                 (23, 24, 25, 26, 27),
                 (28, 29, 30, 31, 31,),
                 (33, 34, 35, 36,),
                 (37, 38, 39, 40, 41,),
                 (42, 43, 44,),
                 (45, 46, 47, 48,),
                 (49, 50, 51,),
                 (52, 53, 54,),
                 (55, 56, 57, 58,),
                 (59, 60, 61,),
                 (62, 63, 64),
                 (65, 66, 67, 68,),
                 (69, 70, 71, 72, 73,),
                 ]

    VSL_groups = [(1, 2, 3, 4,),
                  (5, 6, 7, 8,),
                  (9, 10, 11, 12, 13,),
                  (14, 15, 16, 17,),
                  (18, 19, 20, 21),
                  (22, 23, 24, 25),
                  (26, 27, 28, 29),
                  (30, 31, 32, 33,),
                  (34, 35, 36, 37,),
                  (38, 39, 40,),
                  (41, 42, 43, 44,),
                  (45, 46, 47,),
                  (48, 49, 50,),
                  (51,52,53,54),
                  (55,56,57),
                  (58,59,60),
                  (61,62,63,64),
                  (65,66,67),
                  ]

    ramp_links = [81, 74, 66, 57, 51, 339, 312, 307]

    '''setting of scenarios'''
    # scenarios = [{'group': 16, 'link': 338, 'lane': 3, 'coordinate': 1100},
    #            {'group': 16, 'link': 338, 'lane': 2, 'coordinate': 1100},
    #            {'group': 17, 'link': 310, 'lane': 3, 'coordinate': 1000},]
    scenarios = [{'group': 17, 'link': 306, 'lane': 2, 'coordinate': 5},
                 {'group': 17, 'link': 307, 'lane': 4, 'coordinate': 250},
                 {'group': 17, 'link': 307, 'lane': 3, 'coordinate': 250},]



    Nsec = 18    #number of link groups
    Ninput = 10  #number of vehcle input points

    simResolution = 5   #simulation steps per second
    Tctrl_sec = 30
    Tdata_sec = 5
    #demands = [4500, 280, 180, 400, 650, 10, 50, 180, 100, 50, 90, 280, 550] # daily average vehicle demands
    demands = [4500, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    for i in xrange(len(demands)):
        demands[i] *= demandRatio
    vehCount = 2000    #Simulation stops when so many vehicles get through the network

    car_class = 10
    truck_class = 20
    vf = {car_class: 65, truck_class: 65} #default speed limit of cars and trucks

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
    vslOld = vsl



    '''Define log file names'''
    speedFileName = os.path.join(folderDir, "speedLog_%d.txt"%randomSeed)
    densityFileName = os.path.join(folderDir, "densityLog_%d.txt"%randomSeed)
    flowSectionFileName = os.path.join(folderDir, "flowSectionLog_%d.txt"%randomSeed)
    vslFileName_car = os.path.join(folderDir, "vslLog_car_%d.txt"%randomSeed)
    vslFileName_truck = os.path.join(folderDir, "vslLog_truck_%d.txt"%randomSeed)
    bottleneckDataFileName = os.path.join(folderDir, 'bottleneckData_%d.txt'%randomSeed)

    '''write file headers'''
    writeHeader(speedFileName, "Average Speed of Each Section", Nsec)
    writeHeader(densityFileName, "Density of Each Section", Nsec)
    writeHeader(flowSectionFileName, "Flow Rate of Each Section", Nsec)
    writeHeader(vslFileName_car, "VSL Command for car of Each Section", Nsec)
    writeHeader(vslFileName_truck, "VSL Command for truck of Each Section", Nsec)
    bottleneckDataFile = open(bottleneckDataFileName, 'w')
    bottleneckDataFile.close()


    ProgID = "VISSIM.Vissim"
    
    '''file paths'''
    networkFileName = "I710.inp"
    layoutFileName = "I710.ini"
    #path = networkDir
    #arrivedVehFileName = os.path.join(path, "ArrivedVehs.txt")
    
    ''' Start VISSIM simulation '''
    ''' COM lines'''   

    try :
        
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
        #a = inspect(net)
        #net.print_detail()

        ''' Set vehicle input '''
        vehInputs = net.VehicleInputs
        #inputNum = vehInputs.Count

        for num in xrange(Ninput):
            vehInput = vehInputs.GetVehicleInputByNumber(num+1)
            vehInput.SetAttValue('VOLUME', demands[num])
            vehInput.SetAttValue('TRAFFICCOMPOSITION', demandComposition)
        
        # print "inputNum =", inputNum
        
        # for vehInput in vehInputs:
        #     vehInput.SetAttValue("VOLUME", vehicleInput_veh_hr)

        ''' Set default speed limit '''
        VSLs = net.DesiredSpeedDecisions
        numberSpeedLimit = VSLs.count
        # print "VSLs =", VSLs, "count =", numberSpeedLimit
        
        for VSL in VSLs :
            VSL.SetAttValue1("DESIREDSPEED", car_class, vf[car_class])
            VSL.SetAttValue1("DESIREDSPEED", truck_class, vf[truck_class])



        ''' Get Data collection and link objects '''
        dataCollections = net.DataCollections
        # print "dataCollections =", dataCollections
        
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
        '''Put LC to upstream bottleneck'''
        link = links.GetLinkByNumber(68)
        link.SetAttValue2('LANECLOSED', 1, car_class, True)
        link.SetAttValue2('LANECLOSED', 1, truck_class, True)

        link = links.GetLinkByNumber(10088)
        link.SetAttValue2('LANECLOSED', 1, car_class, True)
        link.SetAttValue2('LANECLOSED', 1, truck_class, True)

        link = links.GetLinkByNumber(66)
        link.SetAttValue2('LANECLOSED', 1, car_class, True)
        link.SetAttValue2('LANECLOSED', 1, truck_class, True)
        link.SetAttValue2('LANECLOSED', 2, car_class, True)
        link.SetAttValue2('LANECLOSED', 2, truck_class, True)





        busNo = 2
        vehArray3 = [0.0] * busNo
        vehArray4 = [0.0] * busNo

        vehs = net.Vehicles


        # start simulation
        print "Client: Starting simulation"
        print "Vehicle Input Ratio:", demandRatio
        print "Scenario:", idxScenario
        print "Controller:", idxController
        print "Lane Closure:", idxLaneClosure, "section"
        print "Random Seed:", randomSeed
            
        simTime_sec = 0 #Simulation Time in seconds

        # Open arrived vehicle ID log, will be closed after simulation
        # arrivedVehFile.open(arrivedVehFileName.c_str(), ios::out);
        
        step = 0
        totVehNum = 0   #count total vehicles get through the network
            
        #while totVehNum < vehCount:
        while simTime_sec < simulationTime_sec:
            sim.RunSingleStep()
            step += 1
            
            if 0 == step%simResolution:
                # Update simulation time
                simTime_sec += 1
                # Collect Data
                if 0 == simTime_sec % Tdata_sec:

                    for num in xrange(Nsec):
                        dataCollection = dataCollections.GetDataCollectionByNumber(num+1)
                        flowSection[num] = dataCollection.GetResult("NVEHICLES", "sum", 0)
                    # end for i, dataCollection in enumerate(dataCollections[:Nsec]):
                    bottleneckData = dataCollections.GetDataCollectionByNumber(19)
                    occBN = dataCollection.GetResult('OCCUPANCYRATE', 'sum', 0)
                    speedBN = dataCollection.GetResult('SPEED', 'MEAN', 0)
                    bottleneckDataFile = open(bottleneckDataFileName, 'a')
                    bottleneckDataFile.write(str(simTime_sec) + '\t' + str(occBN) + '\t' + str(speedBN) + '\n')
                    bottleneckDataFile.close()




                    for i in xrange(Nsec):
                        Nveh = 0
                        dist = 0
                        denomenator = 0
                        for num in link_groups[i]:
                            link = links.GetLinkByNumber(num)
                            den = link.GetSegmentResult("DENSITY", 0, 0, 1, 0)
                            v = link.GetSegmentResult("SPEED",   0, 0, 1, 0)
                            Nveh += den * link_length[num]
                            dist += v * den * link_length[num] #total distance tranveled by all vehicles in one link
                            #denomenator += link_Nlane[num] * link_length[num]
                            denomenator += link_length[num]
                        density[i] = Nveh / denomenator
                        if 0 == Nveh:
                            speed[i] = 0
                        else:
                            speed[i] = dist / Nveh
                    #meanDensity = calcMeanDensity(density, link_groups, link_length)


                    # end for i in xrange(101, 116+1):

                    if startTime_sec <= simTime_sec:
                        totVehNum += ( flowSection[17])
                        #print totVehNum

                    

                    ''' write log files : flowSection, speed, density '''
                    #   open files and indicate time
                    flowSectionFile = open(flowSectionFileName, "a")
                    flowSectionFile.write(str(simTime_sec)); flowSectionFile.write("\t");
                    
                    speedFile = open(speedFileName, "a")
                    speedFile.write(str(simTime_sec)); speedFile.write("\t")

                    densityFile = open(densityFileName, "a")
                    densityFile.write(str(simTime_sec)); densityFile.write("\t")
                        
                    #   write values. Nsec loop
                    for fl, sp, de in zip(flowSection, speed, density):
                        flowSectionFile.write(str(fl)); flowSectionFile.write("\t")
                        speedFile.write(str(sp)); speedFile.write('\t')
                        densityFile.write(str(de)); densityFile.write('\t')
                    # end for fl, sp, de in zip(flowSection, speed, density)

                    # mark end of line and close file
                    flowSectionFile.write('\n'); flowSectionFile.close()
                    speedFile.write('\n'); speedFile.close()
                    densityFile.write('\n'); densityFile.close()
                    # end write log files : flowSection, speed, density


                # end if 0 == simTime_sec % Tdata_sec
                # end Collect Data

                if startTime_sec == simTime_sec:
                    # Set Scenarios
                    for i in xrange(busNo):
                        vehArray3[i] = vehs.AddVehicleAtLinkCoordinate(300, 0, scenarios[idxScenario]['link'], scenarios[idxScenario]['lane'], scenarios[idxScenario]['coordinate'] + 40 * i, 0)
                    LC_Control(scenarios[idxScenario], links, link_groups, ramp_links, 2000*idxLaneClosure)

                if endTime_sec == simTime_sec:
                    # Remove the incident
                    for veh in vehArray3:
                        vehs.RemoveVehicle(veh.ID)
                    for link in links:
                        ID = link.ID
                        # Open all lanes
                        for lane in range(link_Nlane[ID]):
                            link.SetAttValue2('LANECLOSED', lane+1, car_class, False)
                            link.SetAttValue2('LANECLOSED', lane+1, truck_class, False)
                    '''Put LC to upstream bottleneck'''
                    link = links.GetLinkByNumber(68)
                    link.SetAttValue2('LANECLOSED', 1, car_class, True)
                    link.SetAttValue2('LANECLOSED', 1, truck_class, True)

                    link = links.GetLinkByNumber(10088)
                    link.SetAttValue2('LANECLOSED', 1, car_class, True)
                    link.SetAttValue2('LANECLOSED', 1, truck_class, True)

                    link = links.GetLinkByNumber(66)
                    link.SetAttValue2('LANECLOSED', 1, car_class, True)
                    link.SetAttValue2('LANECLOSED', 1, truck_class, True)
                    link.SetAttValue2('LANECLOSED', 2, car_class, True)
                    link.SetAttValue2('LANECLOSED', 2, truck_class, True)
                    
                if 0 == simTime_sec % Tctrl_sec:
                    if (startTime_sec <= simTime_sec < endTime_sec):
                        #Apply VSL controller
                        vsl = control(scenarios[idxScenario], idxController, idxLaneClosure, density, vsl)
                        pass
                    else:
                        for i in range(Nsec):
                            vsl[i] = dict(vf)


                    for i in range(Nsec):
                        for j in VSL_groups[i]:
                            VSL = VSLs.GetDesiredSpeedDecisionByNumber(j)
                            VSL.SetAttValue1('DESIREDSPEED', car_class, vsl[i][car_class])
                            VSL.SetAttValue1('DESIREDSPEED', truck_class, vsl[i][truck_class])
                    vslFile_car = open(vslFileName_car, 'a')
                    vslFile_truck = open(vslFileName_truck, 'a')
                    vslFile_car.write(str(simTime_sec)); vslFile_car.write('\t')
                    vslFile_truck.write(str(simTime_sec)); vslFile_truck.write('\t')
                    for vsli in vsl:
                        vslFile_car.write(str(vsli[car_class]) + '\t')
                        vslFile_truck.write(str(vsli[truck_class]) + '\t')
                    vslFile_car.write('\n')
                    vslFile_truck.write('\n')
                    vslFile_car.close()
                    vslFile_truck.close()

                    #vslOld = vsl[:]
                    #densityOld = density[:]
                    #meanDensityOld = meanDensity[:]
                    # for i in xrange(len(vsl)):
                    #   vslOld[i] = vsl[i]
                    # for i in xrange(len(density)):
                    #   densityOld[i] = density[i]

    

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

def calcMeanDensity(density, link_groups, link_length):
    '''
    calculate the mean downstream density from density
    density: list of density of each section
    link_groups: list of link groups
    link_length:list of length of each link
    '''
    meanDensity = [0.0] * (len(density) - 1)
    sectionLength = [0.0] * len(density)
    for isec in xrange(len(link_groups)):
        for jlink in link_groups[isec]:
            sectionLength[isec] += link_length[jlink]
    for isec in xrange(len(meanDensity)):
        numVeh = 0
        length = 0
        for i in range(isec+1, len(density)):
            numVeh += density[i] * sectionLength[i]
            length += sectionLength[i]
        meanDensity[isec] = numVeh / length
    return meanDensity


def control(scenario, idxController, idxLaneClosure, density, vsl):
    '''
    scenario: dict of current scenario
    idxController: index of the VSL controller
    idxLaneClosure: index of the Lane Change Controller
    meanDensity: list of mean density of current time
    meanDensityOld: list of mean density of previous time step
    vsl: list of vsl conmmands at current time step
    vslOld: list of vsl commands at previous time step
    '''
    dischargeSec = int(idxLaneClosure)
    dischargeSec = 1
    dischargeLength = [1726.75, 2149.22];

    car_class = 10
    truck_class = 20
    acc_section = scenario['group']
    # Number of sections controlled by the combined LC and VSL controller
    Nsec_controlled = 10 - dischargeSec + 1 
    # Start and end section of VSL control (discharging section not included)
    startSection = acc_section - 9
    endSection = acc_section - dischargeSec

    


    vf = {car_class: 65, truck_class: 65}
    vslMIN = {car_class: 10, truck_class: 10}
    vslMAX = {car_class: 65, truck_class: 65}
    #vslTemp = dict(vsl)

    # Set desired equilibrium point
    wb = 40
    rho_e = [120.0] * Nsec_controlled
    rho_e[0] = 430.0
    #rho_e[-3] = rho_e[-4] = 100
    ve = [vf[car_class]] * (Nsec_controlled - 1)
    #ve[0] = vf[car_class] * rho_e[-1] / rho_e[0]
    ve[0] = 15

    rho = density[startSection: startSection + Nsec_controlled]

    # Calculate the mean density of discharging section
    if dischargeSec == 2:
        disDensity = (density[acc_section-2] * dischargeLength[0] + density[acc_section-1] * dischargeLength[1]) / sum(dischargeLength)
        rho[-1] = disDensity
    
    for i in range(len(rho)):
        rho[i] = max(rho[i], 10)

    e = [0.0] * Nsec_controlled
    for i in range(len(e)):
        e[i] = rho[i] - rho_e[i]

    Lambda = [20] * (Nsec_controlled - 1)

    if 0 == idxController:
        for i in range(len(vsl)):
            vsl[i] = dict(vf)
    else:
        u = [0.0] * (Nsec_controlled - 1)
        for i in range(Nsec_controlled-2):
            u[i] = 1 / rho[i] * (-ve[i] * e[i] - Lambda[i] * e[i+1])
        if e[Nsec_controlled - 1] <= 0:
            u[Nsec_controlled - 2] = 1 / rho[Nsec_controlled - 2] * (-Lambda[Nsec_controlled - 2] * e[Nsec_controlled  - 1] - ve[Nsec_controlled-2] * e[Nsec_controlled-2] + vf[car_class] * e[Nsec_controlled-1])
        else:
            u[Nsec_controlled - 2] = 1 / rho[Nsec_controlled - 2] * (-Lambda[Nsec_controlled - 2] * e[Nsec_controlled  - 1] - ve[Nsec_controlled-2] * e[Nsec_controlled-2] - wb * e[Nsec_controlled-1])
        for i in range(startSection, endSection + 1):
            for vehclass in [car_class, truck_class]:
                vslTemp = u[i -startSection] + ve[i - startSection]
                vslTemp = round(vslTemp * 0.2) * 5
                if vslTemp <= vsl[i][vehclass]:
                    vsl[i][vehclass] = max([vslTemp, vsl[i][vehclass] - 10, vslMIN[vehclass]])
                else:
                    vsl[i][vehclass] = min([vslTemp, vsl[i][vehclass] + 10, vslMAX[vehclass]])
        #vsl[-2][car_class] = vsl[-2][truck_class] = 65
        pass




        # Kp = 0
        # Ki = 2
        # xc = 130
        # for i in range(startSection, endSection + 1)[::-1]:
        #     vslTemp = Kp * (meanDensityOld[i] - meanDensity[i]) + Ki * (xc - meanDensity[i])
        #     if vslTemp < -10:
        #         vslTemp = -10
        #     elif vslTemp > 10:
        #         vslTemp = 10
        #     vslTemp = round(vslTemp * 0.2) * 5
        #     vsl[i][car_class] = vslOld[i][car_class] + vslTemp
        #     vsl[i][truck_class] = vslOld[i][truck_class] + vslTemp

        # for i in xrange(startSection, endSection + 1):
        #   meanDensity = 0.0
        #   meanDensityOld = 0.0

        #   for j in xrange(i, acc_section+1):
        #       meanDensity += density[i]
        #       meanDensityOld += densityOld[i]
        #   # for j in xrange(i, Nsec)

        #   meanDensity *= 1.0/(acc_section - i + 1)
        #   meanDensityOld *= 1.0/(acc_section - i + 1)

        #   if meanDensity >= meanDensityOld :
        #       vslTemp = Kv1 * (meanDensityOld - meanDensity)
        #   else:
        #       vslTemp = Kv2 * (meanDensityOld - meanDensity)
        #   # end if meanDensity >= meanDensityOld

        #   if -10 > vslTemp :
        #       vslTemp = -10.0


        #   vslTemp = round(vslTemp*0.2)*5
        #   vsl[i][car_class] = vslOld[i][car_class] + vslTemp
        #   vsl[i][truck_class] = vslOld[i][truck_class] + vslTemp

            # for vehclass in [car_class, truck_class]:
            #     vsl[i][vehclass] = max(vsl[i][vehclass], vslMIN[vehclass])
            #     vsl[i][vehclass] = min(vsl[i][vehclass], vslMAX[vehclass], vsl[i+1][vehclass] + 10)
            # pass
    return vsl

            # end if vslMIN > vsl[i]

        # end for i in xrange(startSection - 1, endSection)


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


#same as main()
if "__main__" == __name__:
    
    scenario = 0
    nMonteCarlo = 5           #Number of Monte Carlo
    #inc = [10, 30, 60]
    inc = [60]
    incTime = {10: (1500, 2100), 30: (1500, 3300), 60: (1500, 4800)}
    #ctrl = [(0, 0.1), (1, 2)]
    #ctrl = [(1, 0.1), (0, 2)]
    ctrl = [(0, 2)]
    demandRatio = 0.11 #The demand is 4500veh/hr when demandRatio == 1
    simulationTime_sec = 4800
    randomSeeds = calcRandomSeeds(nMonteCarlo)
    # Vehicle Composition ID
    # 1: 30% Trucks
    # 2: 15% Trucks
    demandComposition = 2

    networkDir = os.path.abspath(os.curdir)

    for iInc in inc:
        startTime_sec, endTime_sec = incTime[iInc]
        for jController, kLaneClosure in ctrl:
            folderDir = os.path.join(networkDir, 'data', 'incident%dmin' %iInc, 'Ctrl%d' %jController, "LC%d" %kLaneClosure)
            mkdir(folderDir)
            for lMonte in randomSeeds:
                runSimulation(simulationTime_sec, startTime_sec, endTime_sec, demandRatio, demandComposition, scenario, jController, kLaneClosure, lMonte, folderDir, networkDir)
                shutil.copyfile(os.path.join(networkDir, 'I710.fzp'), os.path.join(folderDir, 'I710_%d.fzp'%lMonte))

    os.system('pause')

    
    # simulationTime_sec = 3600   #Simulation time (s)
    # startTime_sec = 600       #Time accident starts (s)
    # endTime_sec = 3600         #Time accident ends (s)
    # #vehicleInput_veh_hr = 9000  #vehicle input (veh/hr)
    # #scenario = range(nScenario)              #Index of simulation scenarios
    # scenario = [0]
    # controller = [1]      #Controller Index
    # laneClosure = [2]    #Whether Lane Change control is added
    # #randomSeeds = [16]
    # randomSeeds = calcRandomSeeds(nMonteCarlo)
    
    # nController = len(controller)
    # nLaneClosure = len(laneClosure)
    # nScenario = len(scenario)               #number of simulation scenarios
    # demandRatio = 1.333

    # # Vehicle Composition ID
    # # 1: 30% Trucks
    # # 2: 15% Trucks
    # demandComposition = 2
    
    # absolute path of current directory
   
    # 'reference to' dataDir
    # make relative path \data\input[vehicle input rate]
    #dirBuf = os.path.join("data", "input%d"%demandRatio)
    # join two paths and make a new string
    # dataDir will point to this new string
    #dataDir = os.path.join(dataDir, dirBuf)
    # following two lines could replace lines above
    # networkDir = os.path.abspath(os.curdir)
    # dataDir = os.path.join( networkDir, "data", "input%d"%vehicleInput_veh_hr) )
    
    
    #randomSeeds = [16]
    
    # for iScenario in scenario :
    #     for jController, kLaneClosure in zip(controller, laneClosure):
    #         folderDir = os.path.join(networkDir, "data", "scenario%d"%iScenario, "Ctrl%d"%jController, "LC%d"%kLaneClosure)
    #         mkdir(folderDir)
    #         carFile = open(os.path.join(folderDir, 'carRecord.txt'), 'w')
    #         truckFile = open(os.path.join(folderDir, 'truckRecord.txt'), 'w')
    #         carSum = []
    #         truckSum = []
    #         for lMonte in randomSeeds:
    #             runSimulation(simulationTime_sec, startTime_sec, endTime_sec, demandRatio, demandComposition, iScenario, jController, kLaneClosure, lMonte, folderDir, networkDir)
    #             shutil.copyfile(os.path.join(networkDir, 'I710.fzp'), os.path.join(folderDir, 'I710_%d.fzp'%lMonte))
    #         #   print "Evaluating Simulation Results"
    #         #   [carTotal, truckTotal] = fzp.fzpEvaluation()
    #         #   carFile.write(str(carTotal) + '\n')
    #         #   truckFile.write(str(truckTotal) + '\n')
    #         #   carSum.append(carTotal)
    #         #   truckSum.append(truckTotal)
    #         #   os.system("cls")
    #         # carSum = np.array(carSum).mean(axis = 0)
    #         # truckSum = np.array(truckSum).mean(axis = 0)
    #         # carFile.write(str(carSum) + '\n')
    #         # truckFile.write(str(truckSum) + '\n')

    #         # carFile.close()
    #         # truckFile.close()
    #         pass

    # #email_interface.mail(DEST, getProgramName(), "program finished running") 
    # #win_interface.shutdown()
    # os.system("pause")
