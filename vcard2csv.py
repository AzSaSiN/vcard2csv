#!/usr/bin/env python3

from time import pthread_getcpuclockid
import vobject
import glob
import csv
import argparse
import os.path
import sys
import logging
import collections
import re

vcard_dir=os.path.join(os.path.dirname(os.path.abspath(__file__)), "Contacts", "test")

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
    '(',
    ')',
    '-',
    ' ',
    ':',
    ','
    ]

prefix_dir = [
    'E'
    'N',
    'NE',
    'NW',
    'S',
    'SE',
    'SW',
    'W',
    ]

google_out_headers = [
    'Name',
    'Given Name',
    'Additional Name',
    'Family Name',
    'Yomi Name',
    'Given Name Yomi',
    'Additional Name Yomi',
    'Family Name Yomi',
    'Name Prefix',
    'Name Suffix',
    'Initials',
    'Nickname',
    'Short Name',
    'Maiden Name',
    'Birthday',
    'Gender',
    'Location',
    'Billing Information',
    'Directory Server',
    'Mileage',
    'Occupation',
    'Hobby',
    'Sensitivity',
    'Priority',
    'Subject',
    'Notes',
    'Language',
    'Photo',
    'Group Membership',
    'E-mail 1 - Type',
    'E-mail 1 - Value',
    'E-mail 2 - Type',
    'E-mail 2 - Value',
    'Phone 1 - Type',
    'Phone 1 - Value',
    'Phone 2 - Type',
    'Phone 2 - Value',
    'Phone 3 - Type',
    'Phone 3 - Value',
    'Phone 4 - Type',
    'Phone 4 - Value',
    'Address 1 - Type',
    'Address 1 - Formatted',
    'Address 1 - Street',
    'Address 1 - City',
    'Address 1 - PO Box',
    'Address 1 - Region',
    'Address 1 - Postal Code',
    'Address 1 - Country',
    'Address 1 - Extended Address',
    'Address 2 - Type',
    'Address 2 - Formatted',
    'Address 2 - Street',
    'Address 2 - City',
    'Address 2 - PO Box',
    'Address 2 - Region',
    'Address 2 - Postal Code',
    'Address 2 - Country',
    'Address 2 - Extended Address',
    'Organization 1 - Type',
    'Organization 1 - Name',
    'Organization 1 - Yomi Name',
    'Organization 1 - Title',
    'Organization 1 - Department',
    'Organization 1 - Symbol',
    'Organization 1 - Location',
    'Organization 1 - Job Description'
    ]

road_list = [
    'ALY',  'AVE', 'BLVD','CSWY', 'CTR',
    'CIR',  'CT',  'CV',  'XING', 'DR',
    'EXPY', 'EXT', 'FWY', 'GRV',  'HWY',
    'HOLW', 'JCT', 'LN',  'MTWY', 'OPAS',
    'PARK', 'PKWY','PL',  'PLZ',  'PT',
    'RD',   'RTE', 'SKWY','SQ',   'ST',
    'TER',  'TRL', 'WAY'
    ]

road_abrr = {
    'ALY' : 'Alley',
    'AVE' : 'Avenue',
    'BLVD': 'Boulevard',
    'CSWY': 'Causeway',
    'CTR' : 'Center',
    'CTR' : 'Circle',
    'CT'  : 'Court',
    'CV'  : 'Cove',
    'XING': 'Crossing',
    'DR'  : 'Drive',
    'EXPY': 'Expressway',
    'EXT' : 'Extension',
    'FWY' : 'Freeway',
    'GRV' : 'Grove',
    'HWY' : 'Highway',
    'HOLW': 'Hollow',
    'JCT' : 'Junction',
    'LN'  : 'Lane',
    'MTWY': 'Motorway',
    'OPAS': 'Overpass',
    'PARK': 'Park',
    'PKWY': 'Parkway',
    'PL'  : 'Place',
    'PLZ' : 'Plaza',
    'PT'  : 'Point',
    'RD'  : 'Road',
    'RTE' : 'Route',
    'SKWY': 'Skyway',
    'SQ'  : 'Square',
    'ST'  : 'Street',
    'TER' : 'Terrace',
    'TRL' : 'Trail',
    'WAY' : 'Way'
}

states_list = [ 
    'AK', 'AL', 'AR', 'AZ', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA',
    'HI', 'IA', 'ID', 'IL', 'IN', 'KS', 'KY', 'LA', 'MA', 'MD', 'ME',
    'MI', 'MN', 'MO', 'MS', 'MT', 'NC', 'ND', 'NE', 'NH', 'NJ', 'NM',
    'NV', 'NY', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX',
    'UT', 'VA', 'VT', 'WA', 'WI', 'WV', 'WY'
    ]

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

def strip_garbage(item, gc=gc):
    for i in gc:
        if i in item:
            item = item.replace(i, '')
    return item

def get_phone_numbers(vCard):
    cell = home = work = other = fax = None
    for tel in vCard.tel_list:
        if vCard.version.value == '2.1':
            if 'CELL' in tel.singletonparams or 'CELL' in tel.singletonparams:
                cell = strip_garbage(str(tel.value).strip())
            elif 'WORK' in tel.singletonparams:
                work = strip_garbage(str(tel.value).strip())
            elif 'HOME' in tel.singletonparams or 'MAIN' in tel.singletonparams:
                home = strip_garbage(str(tel.value).strip())
            elif 'OTHER' in tel.singletonparams:
                other = strip_garbage(str(tel.value).strip())
            else:
                logging.warning("Warning: Unrecognized phone number category in `{}'".format(vCard))
                tel.prettyPrint()
        elif vCard.version.value == '3.0':
            if 'CELL' in tel.params['TYPE'] or 'VOICE' in tel.params['TYPE']:
                cell = strip_garbage(str(tel.value).strip())
            elif 'HOME' in tel.params['TYPE']:
                home = strip_garbage(str(tel.value).strip())
            elif 'WORK' in tel.params['TYPE'] or 'MAIN' in tel.params['TYPE']:
                work = strip_garbage(str(tel.value).strip())
            elif 'FAX' in tel.params['TYPE']:
                fax = strip_garbage(str(tel.value).strip())
            elif 'OTHER' in tel.params['TYPE']:
                other = strip_garbage(str(tel.value).strip())
            else:
                logging.warning("Unrecognized phone number category in '{}'".format(vCard))
                tel.prettyPrint()
        else:
            raise NotImplementedError("Version not implemented: '{}'".format(vCard.version.value))
    return cell, home, work, fax, other

def get_phone_from_notes(vCard):
    if vCard:
        for item in vCard:
            if 'Phone' in item:
                phone = item.split(' ')
                if len(phone[-1]) == 8 and len(phone[-2]) == 5:
                    type = strip_garbage(phone[2]).lower()
                    number = strip_garbage(phone[-2] + phone[-1]) 
    return type, number


def get_address(vCard):
    '''
    Unstandardized street 	Standardized street
	                        Pre  Street   Suffix Post
    NORTHEAST 1ST WAY 		NE   1st       Way
    SO 234 WEST 		    S    234             W
    FIFTH AVENUE 			     5th       Ave
    SO. FILBERT LANE 		S 	 Filbert   Ln
    HWY CONTRACT 2 			     HC 2
    LUCINDA STR. NE 		     Lucinda   St 	NE
    NORTH MAIN ST EAST 		N 	 Main      St 	E
    P.O. BOX 5 			         PO Box
    RURAL RTE 3 BOX 10 		     RR 3
    SUTTER PLACE 			     Sutter    Pl
    '''
    home = work = other = None
    adr_type = street = city = po_box = region = postal_code = country = ext = None
    for adr in vCard.adr_list:
        if vCard.version.value == '2.1':
            print("Card v2")
            # if 'WORK' in adr.singletonparams:
            #     work = str(adr.value).strip()
            # elif 'HOME' in adr.singletonparams:
            #     home = str(adr.value).strip()
            # elif 'OTHER' in adr.singletonparams:
            #     other = str(adr.value).strip()
            # else:
            #     logging.warning("Warning: Unrecognized address category in `{}'".format(vCard))
            #     adr.prettyPrint()
        elif vCard.version.value == '3.0':

            print("\n--- START ADDRESS ---\n",adr, "\n--- END ADDRESS ---\n")
            if 'WORK' in adr.params['TYPE']:
                out = str(adr.value).strip()
                print("--- OUT ---", out)
            elif 'WORK' in adr.params['TYPE']:
                out = str(adr.value).strip()
                print("--- OUT ---", out)
            elif 'WORK' in adr.params['TYPE']:
                out = str(adr.value).strip()
                print("--- OUT ---", out)
            else:
                print("--- ### GET TO THe CHOPPA ### ---")
            #     for i in gc:
            #         if i in work:
            #              work = work.replace(i, ' ')
            #     work = work.split()
            #     address = []
            #     value = ''
            #     for i in work:
            #         # print(i)
            #         value += i + ' '
            #         if i.upper() in road_types: # Street
            #             address.append(value.strip().title())
            #             value = ''
            #         elif len(address) == 1: # City
            #             address.append(value.strip().capitalize())
            #             value = ''
            #         elif i.upper() in states_list: # State
            #             address.append(value.strip().upper())
            #             value = ''
            #         elif len(address) == 3: # ZipCode
            #             address.append(value.strip().upper())
            #             value = ''
            #         elif len(address) == 4 and i == work[len(work[:-1])]: # Country
            #             address.append(value.strip().title())
            #             value = ''
            #     # work = ','.join(address)
            #     work = address
            #     print("\n>work:", work)

            # elif 'HOME' in adr.params['TYPE']:
            #     home = str(adr.value).strip()
            #     for i in gc:
            #         if i in home:
            #             home = home.replace(i, ' ')
            #     home = home.split()
            #     address = []
            #     value = ''
            #     for i in home:
            #         value += i + ' '
            #         if i.upper() in road_types: # Street
            #             address.append(value.strip().title())
            #             value = ''
            #         elif len(address) == 1: # City
            #             address.append(value.strip().capitalize())
            #             value = ''
            #         elif i.upper() in states_list: # State
            #             address.append(value.strip().upper())
            #             value = ''
            #         elif len(address) == 3: # ZipCode
            #             address.append(value.strip().upper())
            #             value = ''
            #         elif len(address) == 4 and i == home[len(home[:-1])]: # Country
            #             address.append(value.strip().title())
            #             value = ''
            #     # home = ','.join(address)
            #     home = address
            #     print("\n>home:", work)
            # elif 'OTHER' in adr.params['TYPE']:
            #     other = str(adr.value).strip()
            #     for i in gc:
            #         if i in other:
            #              other = other.replace(i, ' ')
            #     other = other.split()
            #     address = []
            #     value = ''
            #     for i in other:
            #         value += i + ' '
            #         if i.upper() in road_types: # Street
            #             address.append(value.strip().title())
            #             value = ''
            #         elif len(address) == 1: # City
            #             address.append(value.strip().capitalize())
            #             value = ''
            #         elif i.upper() in states_list: # State
            #             address.append(value.strip().upper())
            #             value = ''
            #         elif len(address) == 3: # ZipCode
            #             address.append(value.strip().upper())
            #             value = ''
            #         elif len(address) == 4 and i == other[len(other[:-1])]: # Country
            #             address.append(value.strip().title())
            #             value = ''
            #     # other = ','.join(address)
            #     other = address
            #     print("\n>other:", work)

            # else:
            #     logging.warning("Unrecognized address category in `{}'".format(vCard))
            #     adr.prettyPrint()
        else:
            raise NotImplementedError("Version not implemented: {}".format(vCard.version.value))
    # print(home, work, other)
    # return home, work, other
    return adr_type, street, city, po_box, region, postal_code, country, ext


def get_info_list(vCard, vcard_filepath):
    def check_key(dict, key):
        ''' Check if there is empty slipt availabe 
        and the phone number is notin vcard already
        '''
        if key:
            if key not in dict.keys():
                return True
        else:
            return False

    def fill_cards(phone_type, phone_num, pos):
        ''' Fill Cards with phone type and phone number
        0 - Mobile
        1 - Worke 
        2 - Home
        3 - Fax
        4 - Other
        '''
        if phone_type == 0:
            vcard['Phone ' + pos + ' - Type'] = 'Mobile'
            vcard['Phone ' + pos + ' - Value'] = phone_num
        elif phone_type == 1:
            vcard['Phone ' + pos + ' - Type'] = 'Work'
            vcard['Phone ' + pos + ' - Value'] = phone_num
        elif phone_type == 2:
            vcard['Phone ' + pos + ' - Type'] = 'Home'
            vcard['Phone ' + pos + ' - Value'] = phone_num
        elif phone_type == 3:
            vcard['Phone ' + pos + ' - Type'] = 'Fax'
            vcard['Phone ' + pos + ' - Value'] = phone_num
        elif phone_type == 4:
            vcard['Phone ' + pos + ' - Type'] = 'Other'
            vcard['Phone ' + pos + ' - Value'] = phone_num

    vcard = collections.OrderedDict()
    for column in google_out_headers:
        vcard[column] = None
    name = cell = work = home = fax = other = email = note = None
    vCard.validate()
    for key, val in list(vCard.contents.items()):
        # print(key)
        if key == 'fn':
            vcard['Given Name'] = vCard.fn.value
        elif key == 'n':
            name = str(vCard.n.valueRepr()).replace('  ', ' ').strip()
            vcard['Name'] = name
        elif key == 'tel':
            try:
                cell, home, work, fax, other = get_phone_numbers(vCard)
            except Exception as err:
                logging.error("Error reading card '{}', with Error: {}".format(vCard, err))
            a = [cell, work, home, fax, other]
            for i in range(len(a)):
                if not check_key(vcard, vcard['Phone 1 - Value']):
                    fill_cards(i, a[i], str(1))
                elif not check_key(vcard, vcard['Phone 2 - Value']):
                    fill_cards(i, a[i], str(2))
                elif not check_key(vcard, vcard['Phone 3 - Value']):
                    fill_cards(i, a[i], str(3))
                elif not check_key(vcard, vcard['Phone 4 - Value']):
                    fill_cards(i, a[i], str(4))
        elif key == 'email':
            email = str(vCard.email.value).strip()
            vcard['E-mail 1 - Type'] = "Work"
            vcard['E-mail 1 - Value'] = email.lower()
        elif key == 'adr':
            # adrr_type = street = city = po_box =region = postal_code = country = ext = get_address(vCard)
            # home, work, other = get_address(vCard)
            # vcard['Address 1 - Formatted'] = home
            # vcard['Address 1 - Formatted'] = work
            # vcard['Address 1 - Formatted'] = other
            print(get_address(vCard))




















        elif key == 'nickname':
            nickname = str(vCard.nickname.value)
            vcard['Nickname'] = nickname
            print("\n****** NICKNAME *****\n*****", nickname,"*****\n")
        elif key == 'note':
            note = str(vCard.note.value)
            vcard['Notes'] = note
        elif key == 'title':
            title = str(vCard.title.value)
            vcard['Name Prefix'] = title
            # print("TITLE:", title)
        else:
            # An unused key, like `adr`, `title`, `url`, etc.
            # print("***UNSUSED KEY***", key, "\t", val)
            
            pass
    if name is None:
        logging.warning("no name for vCard in file `{}'".format(vcard_filepath))
    if all(telephone_number is None for telephone_number in [cell, work, home, fax, other]):
        ''' Last resort to extract phone numbers'''
        try:
            type, number = get_phone_from_notes(vCard.note.value.split('\n'))
            if not check_key(vcard, vcard['Phone 1 - Value']):
                vcard['Phone 1 - Type'] = type.title()
                vcard['Phone 1 - Value'] = number
            elif not check_key(vcard, vcard['Phone 2 - Value']):
                vcard['Phone 2 - Type'] = type.title()
                vcard['Phone 2 - Value'] = number
            elif not check_key(vcard, vcard['Phone 3 - Value']):
                vcard['Phone 3 - Type'] = type.title()
                vcard['Phone 3 - Value'] = number
            elif not check_key(vcard, vcard['Phone 4 - Value']):
                vcard['Phone 4 - Type'] = type.title()
                vcard['Phone 4 - Value'] = number
        except Exception as err:
            pass
            # logging.info("Error reading notes from card '{}'".format(vCard))
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
    parser = argparse.ArgumentParser(description='Convert a bunch of vCard (.vcf) files to a single CSV file.')
    parser.version = '1.0'
    # parser.add_argument('--h',
        # help='Convert a bunch of vCard (.vcf) files to a single CSV file.'
    # )
    # parser = argparse.ArgumentParser(
    #     description='Convert a bunch of vCard (.vcf) files to a single CSV file.'
    # )
    parser.add_argument(
        '-i',
        metavar = '<in path>',
        type=readable_directory,
        help='Directory to read vCard files from.'
    )
    parser.add_argument(
        '-o',
        metavar = "<out file path>",
        type=writable_file,
        help='Output file',
    )
    parser.add_argument(
        '-V',
        '--verbose',
        help='More verbose logging',
        dest="loglevel",
        default=logging.INFO,
        action="store_const",
        const=logging.NOTSET,
    )
    parser.add_argument(
        '-d',
        '--debug',
        help='Enable debugging logs "default INFO" (INFO, LOG, WARN, TRACE, DEBUG, ERROR, FATAL)',
        action="store_const",
        dest="loglevel",
        const=logging.NOTSET,
    )
    opts = parser.parse_args()
    # args = parser.parse_args()
    logging.basicConfig(level=opts.loglevel)
    # print("---- in ", opts.i)
    # print("---- out", opts.o )

    if not opts.i:
        # print("Parser:", opts)
        # parser.print_help()
        opts.i = vcard_dir
        # print("opt.i:", opts.i)

    if opts.i:
        vcard_pattern = os.path.join(opts.i, "*.vcf")
        # print("PATERN:",vcard_pattern)
        vcard_paths = sorted(glob.glob(vcard_pattern))

    # if vcard_pattern:

        if len(vcard_paths) == 0:
            logging.warning("no files ending with `.vcf` in directory `{}'".format(opts.read_dir))
            sys.exit(2)
    # if opts.o:
    # with open(opts.o, 'w') as tsv_fp:
    #     writer = csv.writer(tsv_fp, delimiter=',')
    #     writer.writerow(column_order)

        for vcard_path in vcard_paths:
            for vcard in get_vcards(vcard_path):
                # print("\n--- vCard Info end ---\n")
                vcard_info = get_info_list(vcard, vcard_path)
                # print("--- vCard Info begin ---\n", vcard_info, "\n--- vCard Info begin ---\n")
                # writer.writerow(list(vcard_info.values()))
                # print(list(vcard_info.values()))
        




if __name__ == "__main__":
    main()
