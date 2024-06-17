#!/usr/local/bin/python3
# -*- coding: utf-8 -*-

import logging
import argparse
from pathlib import Path

import Utils

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
            extractdata = Utils.ExtractData(args.email, args.password)
        else:
            if args.startdate <= args.enddate:
                logging.info('Within Range')
                print(args.startdate, args.enddate)
                extractdata = Utils.ExtractData(args.email, args.password, args.startdate, args.enddate)
            else:
                logging.error("Out of Range")
                raise ValueError("startdate must be before or equal to enddate")
    else:
        raise ValueError("Email and password must be provided")

    output_dir = args.output
    output_dir.mkdir(parents=True, exist_ok=True)
    attachment_dir = args.attachments
    attachment_dir.mkdir(parents=True, exist_ok=True)
    extractdata.iterate_emails(output_dir, attachment_dir)

print("Download complete.")
