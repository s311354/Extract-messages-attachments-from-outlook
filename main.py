#!/usr/local/bin/python3
# -*- coding: utf-8 -*-

import logging
import argparse
import sys
from pathlib import Path

import Utils

IMAPserver = "outlook.office365.com"
ImapPort = 993

# Configure logging to write to both file and console
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

# Console handler
log_handler_console = logging.StreamHandler(sys.stdout)
log_handler_console.setFormatter(log_formatter)

# Create a logger
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
logger.addHandler(log_handler_console)

def create_parser():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        '-m', '--email',
        default = 'None',
        type = str,
        help = 'The owner you want to set the email user for credentials.'
    )
    parser.add_argument(
        '-p', '--password',
        default = 'None',
        type = str,
        help = 'The owner you want to set the email password for credentials.'
    )
    parser.add_argument(
        '-s', '--startdate',
        default = None,
        type = str,
        help = 'The owner you want to define the start date to extract the email content. Ex. 01-Jan-2024'
    )
    parser.add_argument(
        '-e', '--enddate',
        default = None,
        type = str,
        help = 'The owner you want to define the end date to extract the email content., Ex. 01-Jan-2024'
    )
    parser.add_argument(
        '-o', '--output',
        default = Path.cwd() / "Output",
        type = str,
        help = 'The owner you want to choose the folder to save emails.'
    )
    parser.add_argument(
        '-a', '--attachments',
        default = Path.cwd() / "Attachments",
        type = str,
        help = 'The owner you want to choose the folder to save emails.'
    )

    return parser

if __name__ == '__main__':
    # Get arguments
    parser = create_parser()
    args = parser.parse_args()

    if args.email and args.password:

        if args.startdate is None or args.enddate is None:
            extractdata = Utils.ExtractData(args.email, args.password, ImapPort)
        else:
            if args.startdate <= args.enddate:
                # print(args.startdate, args.enddate)
                extractdata = Utils.ExtractData(IMAPserver, args.email, args.password, ImapPort, args.startdate, args.enddate)
                logging.info(f"Connected to the email server {IMAPserver}.")
            else:
                logging.error(f"LOGIN failed (Retry in setting up startdate/enddatey ...)")
                sys.exit(1)
    else:
        logging.error(f"Email and password must be provided")
        sys.exit(1)

    output_dir = args.output
    output_dir.mkdir(parents=True, exist_ok=True)
    attachment_dir = args.attachments
    attachment_dir.mkdir(parents=True, exist_ok=True)
    extractdata.iterate_emails(output_dir, attachment_dir)

logging.info(f"Download complete.")
sys.exit(0)