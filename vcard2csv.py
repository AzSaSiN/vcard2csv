#!/usr/bin/env python3

import vobject
import glob
import csv
import argparse
import os.path
import sys
import logging
import collections
import re

gc = [
    '\n', 
    '\\',
    '&amp ', 
    '&amp;',
    'amp;',
    'amp ,', 
    '&AMP ,',
    '&BR&,',
    '<',
    '%',
    '  ', 
    'GT;',
    'gt;',
    'LT;',
    'lt;',
    '.br',
    'br',
    '&',
    ' , ',
    '.',
    ','
]

google_out_headers = [
    "Name",
    "Given Name",
    "Additional Name",
    "Family Name",
    "Yomi Name",
    "Given Name Yomi",
    "Additional Name Yomi",
    "Family Name Yomi",
    "Name Prefix",
    "Name Suffix",
    "Initials",
    "Nickname",
    "Short Name",
    "Maiden Name",
    "Birthday",
    "Gender",
    "Location",
    "Billing Information",
    "Directory Server",
    "Mileage",
    "Occupation",
    "Hobby",
    "Sensitivity",
    "Priority",
    "Subject",
    "Notes",
    "Language",
    "Photo",
    "Group Membership",
    "E-mail 1 - Type",
    "E-mail 1 - Value",
    "E-mail 2 - Type",
    "E-mail 2 - Value",
    "Phone 1 - Type",
    "Phone 1 - Value",
    "Phone 2 - Type",
    "Phone 2 - Value",
    "Phone 3 - Type",
    "Phone 3 - Value",
    "Phone 4 - Type",
    "Phone 4 - Value",
    "Address 1 - Type",
    "Address 1 - Formatted",
    "Address 1 - Street",
    "Address 1 - City",
    "Address 1 - PO Box",
    "Address 1 - Region",
    "Address 1 - Postal Code",
    "Address 1 - Country",
    "Address 1 - Extended Address",
    "Address 2 - Type",
    "Address 2 - Formatted",
    "Address 2 - Street",
    "Address 2 - City",
    "Address 2 - PO Box",
    "Address 2 - Region",
    "Address 2 - Postal Code",
    "Address 2 - Country",
    "Address 2 - Extended Address",
    "Organization 1 - Type",
    "Organization 1 - Name",
    "Organization 1 - Yomi Name",
    "Organization 1 - Title",
    "Organization 1 - Department",
    "Organization 1 - Symbol",
    "Organization 1 - Location",
    "Organization 1 - Job Description"
]

road_types = ["RD", "ST", "AVE", "BLVD", "LN",
            "DR", "WAY", "CT", "PLZ", "TER",
            "PL", "FR", "HWY", "TPKE", "FWY",
            "PKWY", "CSWY", "BLTWY", "CRES",
             "ALY", "ESP", "ROAD", "STREET"]

states_list = [ 'AK', 'AL', 'AR', 'AZ', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA',
           'HI', 'IA', 'ID', 'IL', 'IN', 'KS', 'KY', 'LA', 'MA', 'MD', 'ME',
           'MI', 'MN', 'MO', 'MS', 'MT', 'NC', 'ND', 'NE', 'NH', 'NJ', 'NM',
           'NV', 'NY', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX',
           'UT', 'VA', 'VT', 'WA', 'WI', 'WV', 'WY']

states_abrr = {
    'AK': 'Alaska',
    'AL': 'Alabama',
    'AR': 'Arkansas',
    'AZ': 'Arizona',
    'CA': 'California',
    'CO': 'Colorado',
    'CT': 'Connecticut',
    'DC': 'District of Columbia',
    'DE': 'Delaware',
    'FL': 'Florida',
    'GA': 'Georgia',
    'HI': 'Hawaii',
    'IA': 'Iowa',
    'ID': 'Idaho',
    'IL': 'Illinois',
    'IN': 'Indiana',
    'KS': 'Kansas',
    'KY': 'Kentucky',
    'LA': 'Louisiana',
    'MA': 'Massachusetts',
    'MD': 'Maryland',
    'ME': 'Maine',
    'MI': 'Michigan',
    'MN': 'Minnesota',
    'MO': 'Missouri',
    'MS': 'Mississippi',
    'MT': 'Montana',
    'NC': 'North Carolina',
    'ND': 'North Dakota',
    'NE': 'Nebraska',
    'NH': 'New Hampshire',
    'NJ': 'New Jersey',
    'NM': 'New Mexico',
    'NV': 'Nevada',
    'NY': 'New York',
    'OH': 'Ohio',
    'OK': 'Oklahoma',
    'OR': 'Oregon',
    'PA': 'Pennsylvania',
    'RI': 'Rhode Island',
    'SC': 'South Carolina',
    'SD': 'South Dakota',
    'TN': 'Tennessee',
    'TX': 'Texas',
    'UT': 'Utah',
    'VA': 'Virginia',
    'VT': 'Vermont',
    'WA': 'Washington',
    'WI': 'Wisconsin',
    'WV': 'West Virginia',
    'WY': 'Wyoming'
}

def get_phone_numbers(vCard):
    cell = home = work = other = None
    for tel in vCard.tel_list:
        if vCard.version.value == '2.1':
            if 'CELL' in tel.singletonparams:
                cell = str(tel.value).strip()
            elif 'WORK' in tel.singletonparams:
                work = str(tel.value).strip()
            elif 'HOME' in tel.singletonparams:
                home = str(tel.value).strip()
            elif 'MAIN' in tel.singletonparams:
                work = str(tel.value).strip()
            elif 'OTHER' in tel.singletonparams:
                other = str(tel.value).strip()
            elif 'pref' in tel.singletonparams:
                other = str(tel.value).strip()
            else:
                logging.warning("Warning: Unrecognized phone number category in `{}'".format(vCard))
                tel.prettyPrint()
        elif vCard.version.value == '3.0':
            if 'CELL' in tel.params['TYPE']:
                cell = str(tel.value).strip()
            elif 'WORK' in tel.params['TYPE']:
                work = str(tel.value).strip()
                print('################', work, '################')
                if not work:
                    print('################ NO ################')
            elif 'HOME' in tel.params['TYPE']:
                home = str(tel.value).strip()
            elif 'MAIN' in tel.params['TYPE']:
                work = str(tel.value).strip()
            elif 'OTHER' in tel.params['TYPE']:
                other = str(tel.value).strip()
            elif 'pref' in tel.params['TYPE']:
                other = str(tel.value).strip()
            else:
                logging.warning("Unrecognized phone number category in `{}'".format(vCard))
                tel.prettyPrint()
        else:
            raise NotImplementedError("Version not implemented: {}".format(vCard.version.value))
    print("CELL:", cell, "\tHOME:", home, "\tWORK:", work, "\tOTHER:", other)
    return cell, home, work, other

def get_address(vCard):
    home = work = other = None
    for adr in vCard.adr_list:
        if vCard.version.value == '2.1':
            if 'WORK' in adr.singletonparams:
                work = str(adr.value).strip()
            elif 'HOME' in adr.singletonparams:
                home = str(adr.value).strip()
            elif 'OTHER' in adr.singletonparams:
                other = str(adr.value).strip()
            else:
                logging.warning("Warning: Unrecognized address category in `{}'".format(vCard))
                adr.prettyPrint()
        elif vCard.version.value == '3.0':
            if 'WORK' in adr.params['TYPE']:
                work = str(adr.value).strip()
                for i in gc:
                    if i in work:
                         work = work.replace(i, ' ')
                work = work.split()
                address = []
                value = ''
                for i in work:
                    value += i + ' '
                    if i.upper() in road_types: # Street
                        address.append(value.strip().title())
                        value = ''
                    elif len(address) == 1: # City
                        address.append(value.strip().capitalize())
                        value = ''
                    elif i.upper() in states_list: # State
                        address.append(value.strip().upper())
                        value = ''
                    elif len(address) == 3: # ZipCode
                        address.append(value.strip().upper())
                        value = ''
                    elif len(address) == 4 and i == work[len(work[:-1])]: # Country
                        address.append(value.strip().title())
                        value = ''
                # work = ','.join(address)
                work = address

            elif 'HOME' in adr.params['TYPE']:
                home = str(adr.value).strip()
                for i in gc:
                    if i in home:
                        home = home.replace(i, ' ')
                home = home.split()
                address = []
                value = ''
                for i in home:
                    value += i + ' '
                    if i.upper() in road_types: # Street
                        address.append(value.strip().title())
                        value = ''
                    elif len(address) == 1: # City
                        address.append(value.strip().capitalize())
                        value = ''
                    elif i.upper() in states_list: # State
                        address.append(value.strip().upper())
                        value = ''
                    elif len(address) == 3: # ZipCode
                        address.append(value.strip().upper())
                        value = ''
                    elif len(address) == 4 and i == home[len(home[:-1])]: # Country
                        address.append(value.strip().title())
                        value = ''
                # home = ','.join(address)
                home = address
            elif 'OTHER' in adr.params['TYPE']:
                other = str(adr.value).strip()
                for i in gc:
                    if i in other:
                         other = other.replace(i, ' ')
                other = other.split()
                address = []
                value = ''
                for i in other:
                    value += i + ' '
                    if i.upper() in road_types: # Street
                        address.append(value.strip().title())
                        value = ''
                    elif len(address) == 1: # City
                        address.append(value.strip().capitalize())
                        value = ''
                    elif i.upper() in states_list: # State
                        address.append(value.strip().upper())
                        value = ''
                    elif len(address) == 3: # ZipCode
                        address.append(value.strip().upper())
                        value = ''
                    elif len(address) == 4 and i == other[len(other[:-1])]: # Country
                        address.append(value.strip().title())
                        value = ''
                # other = ','.join(address)
                other = address

            else:
                logging.warning("Unrecognized address category in `{}'".format(vCard))
                adr.prettyPrint()
        else:
            raise NotImplementedError("Version not implemented: {}".format(vCard.version.value))

    # print(str(adr.value).strip(),"\nhome :", home, "\nwork :", work, "\nother:", other, "\n")
    print("\nhome :", home, "\nwork :", work, "\nother:", other, "\n")


        # "Address 1 - Type",
        # "Address 1 - Formatted",
        # "Address 1 - Street",
        # "Address 1 - City",
        # "Address 1 - PO Box",
        # "Address 1 - Region",
        # "Address 1 - Postal Code",
        # "Address 1 - Country",
        # "Address 1 - Extended Address",
    return home, work, other    


def get_info_list(vCard, vcard_filepath):
    vcard = collections.OrderedDict()
    for column in google_out_headers:
        vcard[column] = None
    name = cell = work = home = other = email = note = None
    vCard.validate()
    for key, val in list(vCard.contents.items()):
        # print("KEY:\t", key, "\tVALUES:\t", val)
        if key == 'fn':
            vcard['Full name'] = vCard.fn.value
            print("Full NAME:", vCard.fn.value)
        elif key == 'n':
            name = str(vCard.n.valueRepr()).replace('  ', ' ').strip()
            vcard['Name'] = name
            print("NAME:", name)
        elif key == 'tel':
            cell, home, work, other = get_phone_numbers(vCard)
            vcard['Cell phone'] = cell
            vcard['Home phone'] = home
            vcard['Work phone'] = work
            vcard['Other phone'] = other
            print ("** CELL:", cell, "\tHOME:", home, "\tWORK:", work, "\tOTHER:", other)
        elif key == 'email':
            email = str(vCard.email.value).strip()
            vcard['Email'] = email
            print("EMAIL:", email)
        elif key == 'adr':
            home, work, other = get_address(vCard)
            vcard['Home phone'] = home
            vcard['Work phone'] = work
            vcard['Work phone'] = other
        elif key == 'nickname':
            nickname = str(vCard.nickname.value)
            vcard['nickname'] = nickname
            print("NICKNAME:", nickname)
        elif key == 'note':
            note = str(vCard.note.value)
            vcard['Note'] = note
        elif key == 'title':
            title = str(vCard.title.value)
            vcard['Note'] = title
            print("TITLE:", title)
        else:
            # An unused key, like `adr`, `title`, `url`, etc.
            # print("***UNSUSED KEY***", key, "\t", val)
            
            pass
    if name is None:
        logging.warning("no name for vCard in file `{}'".format(vcard_filepath))
    if all(telephone_number is None for telephone_number in [cell, work, home, other]):
        try:
            print(vCard.note.value, type(vCard.note.value))
            if 'Phone' in vCard.note.value:
                card =  vCard.note.value.split('\n')
                print(card)
        except:
            pass
        # logging.warning("no telephone numbers for file `{}' with name `{}'".format(vcard_filepath, name))
    return vcard


def get_vcards(vcard_filepath):
    with open(vcard_filepath) as fp:
        all_text = fp.read()
    for vCard in vobject.readComponents(all_text):
        # print("vCard:", vCard)
        yield vCard


def readable_directory(path):
    if not os.path.isdir(path):
        raise argparse.ArgumentTypeError(
            'not an existing directory: {}'.format(path))
    if not os.access(path, os.R_OK):
        raise argparse.ArgumentTypeError(
            'not a readable directory: {}'.format(path))
    return path


def writable_file(path):
    if os.path.exists(path):
        if not os.access(path, os.W_OK):
            raise argparse.ArgumentTypeError(
                'not a writable file: {}'.format(path))
    else:
        # If the file doesn't already exist,
        # the most direct way to tell if it's writable
        # is to try writing to it.
        with open(path, 'w') as fp:
            pass
    return path


def main():
    parser = argparse.ArgumentParser(
        description='Convert a bunch of vCard (.vcf) files to a single TSV file.'
    )
    parser.add_argument(
        'read_dir',
        type=readable_directory,
        help='Directory to read vCard files from.'
    )
    parser.add_argument(
        'tsv_file',
        type=writable_file,
        help='Output file',
    )
    parser.add_argument(
        '-v',
        '--verbose',
        help='More verbose logging',
        dest="loglevel",
        default=logging.WARNING,
        action="store_const",
        const=logging.NOTSET,
    )
    parser.add_argument(
        '-d',
        '--debug',
        help='Enable debugging logs',
        action="store_const",
        dest="loglevel",
        const=logging.DEBUG,
    )
    args = parser.parse_args()
    logging.basicConfig(level=args.loglevel)

    vcard_pattern = os.path.join(args.read_dir, "*.vcf")
    vcard_paths = sorted(glob.glob(vcard_pattern))
    if len(vcard_paths) == 0:
        logging.error("no files ending with `.vcf` in directory `{}'".format(args.read_dir))
        sys.exit(2)

    # with open(args.tsv_file, 'w') as tsv_fp:
    #     writer = csv.writer(tsv_fp, delimiter=',')
    #     writer.writerow(column_order)

    for vcard_path in vcard_paths:
        for vcard in get_vcards(vcard_path):
            # print("\n--- vCard Info end ---\n")
            vcard_info = get_info_list(vcard, vcard_path)
            # print("--- vCard Info begin ---\n", vcard_info)
            # writer.writerow(list(vcard_info.values()))
            # print(list(vcard_info.values()))
    




if __name__ == "__main__":
    main()
    