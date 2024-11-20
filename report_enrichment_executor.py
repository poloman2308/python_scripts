import sys, getopt

from chassis_repot_enriching import \
    chassis_repot_enriching


def resolveargs(argv):
    inputfile = None
    try:
        opts, args = getopt.getopt(argv,"hi",["ifile="])
    except getopt.GetoptError:
        print ('report_enrichment_executor.py -i <inputfile>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ('report_enrichment_executor.py -i <inputfile>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
    return inputfile




if __name__ == "__main__":
   input = resolveargs(sys.argv[1:])
   chassis_repot_enriching(input)