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

def runSimulation(simulationTime_sec, startTime_sec, endTime_sec, idxScenario, idxControler, idxLaneClosure, randonSeed, folderDir, networkDir):
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





                   
                   
                   
                   
                   
                   
                   
                   
                   
                   
                   
                                      
                   


    #Data Collector Groups
    DC_groups = [()

    ]




    Nsec = 18   # number of sections
    simResolution = 5





    
