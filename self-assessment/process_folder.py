#!/usr/bin/env python3

import sys
import os
import logging
import extract_self_assessment


logging.basicConfig(level=logging.DEBUG)


def main():
    if len(sys.argv) != 2:
        print("Usage: {} folder".format(sys.argv[0]), file=sys.stderr)
        sys.exit(1)

    files = [os.path.join(sys.argv[1], f) for f in os.listdir(sys.argv[1]) if os.path.isfile(os.path.join(sys.argv[1], f)) and f.endswith('.xlsx')]

    print("Nume\tPoziție\tEducație\tCercetare\tManagement\tRecunoaștere\tComunitate\tAltele")
    for f in files:
        logging.info("processing {}".format(f))
        try:
            ret = extract_self_assessment.process_file(f)
            print("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}".format(
                ret["name"],
                ret["position"],
                int(ret["education"]),
                int(ret["research"]),
                int(ret["manage"]),
                int(ret["prestige"]),
                int(ret["community"]),
                int(ret["others"])))
        except:
            logging.error("Error processing {}".format(f))


if __name__ == "__main__":
    sys.exit(main())
