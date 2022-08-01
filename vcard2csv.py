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

vcard_dir=os.path.join(os.path.dirname(os.path.abspath(__file__)), "Contacts", "test")

gc = [
    '-',
    ',',
    ':',
    ';',
    '.',
    '.br',
    '(',
    ')',
    '\\',
    '\n', 
    '&',
    '&AMP ,',
    '&amp ', 
    '&amp;',
    '&BR&,',
    '%',
    '<',
    '>',
    ']',
    '[',
    "'",
    'amp ,', 
    'amp;',
    'amp',
    'gt',
    'GT',
    'lt',
    'LT'
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


def clean_up_address (segment):
        ''' Strip garbage charecter from address'''
        card = []
        segment = segment.strip()
        for i in segment.replace('br', ' ').split():
            adr_seg = strip_garbage(i)
            if adr_seg != '':
                card.append(adr_seg)
        return card


def reconstruct_address(addr):
    po_box = street = city = state = postal_code = country = None
    for i in range(len(addr)):
        # PO BOX Construct
        if i == 0 and addr[i].upper() == "PO":
            if addr[1].upper() == "BOX":
                po_box = "PO Box " + addr[2]
            if addr[4].upper() in states_list:
                city = addr[3].title()
                state = addr[4].upper()
                postal_code = addr[5]
        elif i == 1 and addr[i] in prefix_dir:
            # Address Type 'A' Construct
            if addr[5].upper() in states_list:
                street = str(addr[0] + ' ' + str(addr[1].upper()) + ' ' + str(addr[2]) + ' ' + str(addr[3].title()))
                city = addr[4].title()
                state = addr[5].upper()
                postal_code = addr[6]
        elif i == 1 and addr[i].upper() == 'HOME':
            # Address Type 'A' Construct for ER
            if addr[4] in prefix_dir and addr[4] in states_list:
                street = str(addr[3] + ' ' + str(addr[4].upper()) + ' ' + str(addr[5]) + ' ' + str(addr[6].title()))
                city = addr[7].title()
                state = addr[8].upper()
                postal_code = addr[9]
        elif i == 2 and addr[i].upper() in road_list:
            # Address Type 'B' Construct
            if addr[4].upper() in states_list:
                street = str(addr[0]) + ' ' + str(addr[1].title()) + ' ' + str(addr[2].upper())
                city = addr[3].title()
                state = addr[4].upper()
                postal_code = addr[5]
        elif addr[i].lower() == "united" or addr[i] == "US":
            country = "US"
            break
    return po_box, street, city, state, postal_code, country


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


def get_email(vCard):
    home = work = None
    # print("### vCard ###", vCard)
    
    if vCard:
        for email in vCard.email_list:
            if 'HOME' in email.params['TYPE']:
                print(">>>>>", email.value, "HOME")
                home = 'HOME', email.value.lower()
            elif 'WORK' in email.params['TYPE']:
                print(">>>>>", email.value, "WORK") 
                work = 'WORK', email.value.lower()
        # vCard.email.value).strip()
    return home, work

def get_data_from_notes(note):
    phone_type = phone_number = date = work_email = home_email = home_addr = work_addr = other_addr = None
    if note:
        for item in range(len(note)):
            if 'Phone' in note[item]:
                phone = note[item].split(' ')
                if len(phone[-1]) == 8 and len(phone[-2]) == 5:
                    phone_type = strip_garbage(phone[2]).lower()
                    phone_number = strip_garbage(phone[-2] + phone[-1]) 
            if 'Create Date:' in note[item]:
                date = note[item].split(':')[-1].strip()

            if 'Email - Work' in note[item]:
                work_email = 'WORK', note[item].split(':')[-1].lower().strip()
            elif 'Email - Home' in note[item]:
                home_email = 'HOME', note[item].split(':')[-1].lower().strip()

            if 'Address - Home' in note[item]:
                address = clean_up_address(note[item])
                home_addr = reconstruct_address(address)
            elif 'Address - Work' in note[item]:
                address = clean_up_address(note[item])
                work_addr = reconstruct_address(address)
            elif 'Address - Other' in note[item]:
                address = clean_up_address(note[item])
                other_addr = reconstruct_address(address)
    return phone_type, phone_number, date, work_email, home_email, home_addr, work_addr, other_addr

###
def get_notes_from_notes(note):
    for i in note.split('\n'):
        if 'Description:' in i:
            return i[13:]


def get_address(vCard):
    '''
    Exctract and Construct into usable format
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
    for adr in vCard.adr_list:
        if vCard.version.value == '2.1':
            if 'HOME' in adr.singletonparams:
                out = clean_up_address(str(adr.value).strip())
                home = reconstruct_address(out)
            elif 'WORK' in adr.singletonparams:
                out = clean_up_address(str(adr.value).strip())
                work = reconstruct_address(out)
            elif 'OTHER' in adr.singletonparams:
                out = clean_up_address(str(adr.value).strip())
                other = reconstruct_address(out)
            else:
                logging.warning("Warning: Unrecognized address category in `{}'".format(vCard))
                adr.prettyPrint()
        elif vCard.version.value == '3.0':
            if 'HOME' in adr.params['TYPE']:
                out = clean_up_address(str(adr.value).strip())
                home = reconstruct_address(out)
            elif 'WORK' in adr.params['TYPE']:
                out = clean_up_address(str(adr.value).strip())
                work = reconstruct_address(out)
            elif 'OTHER' in adr.params['TYPE']:
                out = clean_up_address(str(adr.value).strip())
                other = reconstruct_address(out)
            else:
                logging.warning("Warning: Unrecognized address category in `{}'".format(vCard))
                adr.prettyPrint()
        else:
            raise NotImplementedError("Version not implemented: {}".format(vCard.version.value))
    return home, work, other


def check_key(dict, key):
        ''' Check if there is empty slipt availabe 
        and the phone number is notin vcard already
        '''
        if key:
            if key not in dict.items():
                return True
        else:
            return False


def fill_phone_cards(phone_type, phone_num, pos, vcard):
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


def fill_address_cards(home, work, other, pos, vcard):
    pos = pos + 1
    if home:
        po_box, street, city, state, postal_code, country = home
        if po_box:
            f_addr = po_box + ', ' + city + ' ' + state + ', ' + str(postal_code) + ', ' + country
            if vcard['Address 1 - Formatted'] != f_addr and vcard['Address 2 - Formatted'] != f_addr:
                vcard['Address ' + str(pos) + ' - Type'] = 'HOME'
                vcard['Address ' + str(pos) + ' - City'] = city
                vcard['Address ' + str(pos) + ' - PO Box'] = po_box
                vcard['Address ' + str(pos) + ' - Region'] = state
                vcard['Address ' + str(pos) + ' - Postal Code'] = postal_code
                vcard['Address ' + str(pos) + ' - Country'] = country
                vcard['Address ' + str(pos) + ' - Formatted'] = f_addr
        else:
            f_addr = street + '. ' + city + ' ' + state + ', ' + str(postal_code) + ' ' + country
            if vcard['Address 1 - Formatted'] != f_addr and vcard['Address 2 - Formatted'] != f_addr:
                vcard['Address ' + str(pos) + ' - Type'] = 'HOME'
                vcard['Address ' + str(pos) + ' - Street'] = street
                vcard['Address ' + str(pos) + ' - City'] = city
                vcard['Address ' + str(pos) + ' - Region'] = state
                vcard['Address ' + str(pos) + ' - Postal Code'] = postal_code
                vcard['Address ' + str(pos) + ' - Country'] = country
                vcard['Address ' + str(pos) + ' - Formatted'] = f_addr
    if work:
        po_box, street, city, state, postal_code, country = work
        if po_box:
            f_addr = po_box + ', ' + city + ' ' + state + ', ' + str(postal_code) + ', ' + country
            if vcard['Address 1 - Formatted'] != f_addr and vcard['Address 2 - Formatted'] != f_addr:
                vcard['Address ' + str(pos) + ' - Type'] = 'WORK'
                vcard['Address ' + str(pos) + ' - City'] = city
                vcard['Address ' + str(pos) + ' - PO Box'] = po_box
                vcard['Address ' + str(pos) + ' - Region'] = state
                vcard['Address ' + str(pos) + ' - Postal Code'] = postal_code
                vcard['Address ' + str(pos) + ' - Country'] = country
                vcard['Address ' + str(pos) + ' - Formatted'] = f_addr
        else:
            f_addr = street + '. ' + city + ' ' + state + ', ' + str(postal_code) + ' ' + country
            if vcard['Address 1 - Formatted'] != f_addr and vcard['Address 2 - Formatted'] != f_addr:
                vcard['Address ' + str(pos) + ' - Type'] = 'WORK'
                vcard['Address ' + str(pos) + ' - Street'] = street
                vcard['Address ' + str(pos) + ' - City'] = city
                vcard['Address ' + str(pos) + ' - Region'] = state
                vcard['Address ' + str(pos) + ' - Postal Code'] = postal_code
                vcard['Address ' + str(pos) + ' - Country'] = country
                vcard['Address ' + str(pos) + ' - Formatted'] = f_addr
    if other:
        po_box, street, city, state, postal_code, country = other
        if po_box:
            f_addr = po_box + ', ' + city + ' ' + state + ', ' + str(postal_code) + ', ' + country
            if vcard['Address 1 - Formatted'] != f_addr and vcard['Address 2 - Formatted'] != f_addr:
                vcard['Address ' + str(pos) + ' - Type'] = 'OTHER'
                vcard['Address ' + str(pos) + ' - City'] = city
                vcard['Address ' + str(pos) + ' - PO Box'] = po_box
                vcard['Address ' + str(pos) + ' - Region'] = state
                vcard['Address ' + str(pos) + ' - Postal Code'] = postal_code
                vcard['Address ' + str(pos) + ' - Country'] = country
                vcard['Address ' + str(pos) + ' - Formatted'] = f_addr
        else:
            f_addr = street + '. ' + city + ' ' + state + ', ' + str(postal_code) + ' ' + country
            if vcard['Address 1 - Formatted'] != f_addr and vcard['Address 2 - Formatted'] != f_addr:
                vcard['Address ' + str(pos) + ' - Type'] = 'OTHER'
                vcard['Address ' + str(pos) + ' - Street'] = street
                vcard['Address ' + str(pos) + ' - City'] = city
                vcard['Address ' + str(pos) + ' - Region'] = state
                vcard['Address ' + str(pos) + ' - Postal Code'] = postal_code
                vcard['Address ' + str(pos) + ' - Country'] = country
                vcard['Address ' + str(pos) + ' - Formatted'] = f_addr
















def print_card(vcard):
    print("*Given Name:  ", vcard['Given Name'])  # PRINT GIVEN NAME
    print("*NAME:        ", vcard['Name'])  # PRINT NAME
    print("*PHONE #1:    ", vcard['Phone 1 - Value']) # PRINT PHONE NUMBER
    print("*PHONE #2:    ", vcard['Phone 2 - Value']) # PRINT PHONE NUMBER
    print("*PHONE #3:    ", vcard['Phone 3 - Value']) # PRINT PHONE NUMBER
    print("*PHONE #4:    ", vcard['Phone 4 - Value']) # PRINT PHONE NUMBER
    print("*EMAL  #1:    ", vcard['E-mail 1 - Value'])
    print("*EMAL  #2:    ", vcard['E-mail 2 - Value'])
    print("*ADDRESS 1:   ", vcard['Address 1 - Formatted']) # PRINT ADDRESS
    print("*ADDRESS 2:   ", vcard['Address 2 - Formatted']) # PRINT ADDRESS
    print("*NICKNAME:    ", vcard['Nickname'])
    print("*TITLE:       ", vcard['Name Prefix'])
    print("*ORGANIZATION:", vcard['Organization 1 - Name'])
    print("*CREATED DATE:", vcard['Created Date'])

    # print("*EM-PHONE:  ",vcard["Phone 1 - Value"])
    # print("*EM-PHONE:  ",vcard["Phone 2 - Value"])
    # print("*EM-PHONE:  ",vcard["Phone 3 - Value"])
    # print("*EM-PHONE:  ",vcard["Phone 4 - Value"])
    print("*NOTE:", vcard['Notes'])


def get_info_list(vCard, vcard_filepath):
    vcard = collections.OrderedDict()

    
            

    print("\n\n\n###################  START  ###################")
    for column in google_out_headers:
        vcard[column] = None
    name = cell = work = home = fax = other = email = note = None
    vCard.validate()
    for key, val in list(vCard.contents.items()):
        # print(list(vCard.contents.items()))
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
            p = [cell, work, home, fax, other]
            for i in range(len(p)):
                if not check_key(vcard, vcard['Phone 1 - Value']):
                    fill_phone_cards(i, p[i], str(1), vcard)
                elif not check_key(vcard, vcard['Phone 2 - Value']):
                    fill_phone_cards(i, p[i], str(2),vcard)
                elif not check_key(vcard, vcard['Phone 3 - Value']):
                    fill_phone_cards(i, p[i], str(3), vcard)
                elif not check_key(vcard, vcard['Phone 4 - Value']):
                    fill_phone_cards(i, p[i], str(4), vcard)
        elif key == 'email':
            # email = str(vCard.email.value).strip()
            # TODO implement two emails list
            # TODO extract email from notes
            home, work = get_email(vCard)
            print(">>>>>>>>>>>>>>>>", home, work, "<<<<<<<<<<<<<<<<<<<")

            for e in (home, work):
                e_type, email = e
                if not check_key(vcard, vcard['E-mail 1 - Value']):
                    if email != vcard['E-mail 1 - Value']:
                        vcard['E-mail 1 - Type'] = e_type
                        vcard['E-mail 1 - Value'] = email
                elif not check_key(vcard, vcard['E-mail 2 - Value']):
                    if email != vcard['E-mail 1 - Value'] and email != vcard['E-mail 2 - Value']:
                        vcard['E-mail 2 - Type'] = e_type
                        vcard['E-mail 2 - Value'] = email.lower()

            
            
        elif key == 'adr':
            home, work, other = get_address(vCard)
            for i in range(3):
                if not check_key(vcard, vcard['Address 1 - Formatted']):
                    fill_address_cards(home, work, other, i, vcard)
                elif not check_key(vcard, vcard['Address 2 - Formatted']):
                    fill_address_cards(home, work, other, i, vcard)
        elif key == 'org':
            org = str(vCard.org.value)
            for i in ('[', ']', "'"):
                org = org.replace(i, '')
            vcard['Organization 1 - Type'] = 'WORK'
            vcard['Organization 1 - Name'] = org
        elif key == 'nickname':
            nickname = str(vCard.nickname.value)
            vcard['Nickname'] = nickname
        elif key == 'note':
            note = str(vCard.note.value)
            vcard['Notes'] = get_notes_from_notes(note)
            
        elif key == 'title':
            title = str(vCard.title.value)
            vcard['Name Prefix'] = title
        else:
            # An unused key, like `adr`, `title`, `url`, etc.
            # print("***UNSUSED KEY***", key, "\t", val)
            em_phone_type, em_phone_number, em_date, em_work_email, em_home_email, em_home_addr, em_work_addr, em_other_addr = get_data_from_notes(vCard.note.value.split('\n'))
            if 
                for e in (em_home_email, em_work_email):
                    e_type, email = e
                    if not check_key(vcard, vcard['E-mail 1 - Value']):
                        if email != vcard['E-mail 1 - Value']:
                            vcard['E-mail 1 - Type'] = e_type
                            vcard['E-mail 1 - Value'] = email
                    elif not check_key(vcard, vcard['E-mail 2 - Value']):
                        if email != vcard['E-mail 1 - Value'] and email != vcard['E-mail 2 - Value']:
                            vcard['E-mail 2 - Type'] = e_type
                            vcard['E-mail 2 - Value'] = email.lower()
            print("EM PHONE TYPE   :", em_phone_type)
            print("EM PHONE NUMBER :", em_phone_number)
            print("EM DATE         :", em_date)
            print("EM EMAIL WORK   :", em_work_email)
            print("EM EMAIL HOME   :", em_home_email)
            print("EM HOME ADDR    :", em_home_addr)
            print("EM WORK ADDR    :", em_work_addr)
            print("EM OTHER ADDR   :", em_other_addr)
            
            # date = get_date_from_note(vCard.note.value.split('\n'))
            vcard['Created Date'] = em_date

            # email = get_email_from_note(vCard.note.value.split('\n'))
            # print("<<EMAIL>>", email)

            pass
    if name is None:
        logging.warning("no name for vCard in file `{}'".format(vcard_filepath))
    if all(telephone_number is None for telephone_number in [cell, work, home, fax, other]):
        ''' Last resort to extract phone numbers'''
        try:
            # type, number = get_phone_from_notes(vCard.note.value.split('\n')) # PHONE NUMER PASS ON
            if not check_key(vcard, vcard['Phone 1 - Value']):
                vcard['Phone 1 - Type'] = em_phone_type.title()
                vcard['Phone 1 - Value'] = em_phone_number
            elif not check_key(vcard, vcard['Phone 2 - Value']):
                vcard['Phone 2 - Type'] = em_phone_type.title()
                vcard['Phone 2 - Value'] = em_phone_number
            elif not check_key(vcard, vcard['Phone 3 - Value']):
                vcard['Phone 3 - Type'] = em_phone_type.title()
                vcard['Phone 3 - Value'] = em_phone_number
            elif not check_key(vcard, vcard['Phone 4 - Value']):
                vcard['Phone 4 - Type'] = em_phone_type.title()
                vcard['Phone 4 - Value'] = em_phone_number
        except Exception as err:
            pass
            # logging.info("Error reading notes from card '{}'".format(vCard))
    if all(Address is None for Address in [work, home, other]):
        # home, work, other = get_address_from_notes(vCard.note.value.split('\n'))
        for i in range(3):
            if not check_key(vcard, vcard['Address 1 - Formatted']):
                fill_address_cards(em_home_addr, em_work_addr, em_other_addr, i, vcard)
            elif not check_key(vcard, vcard['Address 2 - Formatted']):
                fill_address_cards(em_home_addr, em_work_addr, em_other_addr, i, vcard)

    if all(Email is None for Email in [email, em_work_email, em_home_email]):
        email1 = em_work_email
        email2 = em_home_email
        print("((WORK))", email1, "\n((HOME))", email2)


    print_card(vcard)
    print("####################  END  ####################\n\n\n")
    return vcard


def get_vcards(vcard_filepath):
    with open(vcard_filepath) as fp:
        all_text = fp.read()
    for vCard in vobject.readComponents(all_text):
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
    logging.basicConfig(level=opts.loglevel)

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
                # print("--- vCard Info begin ---\n",
                #  vcard_info, "\n--- vCard Info begin ---\n")
                # writer.writerow(list(vcard_info.values()))
                # print(list(vcard_info.values()))
        




if __name__ == "__main__":
    main()
