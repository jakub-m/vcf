#!/usr/bin/env python2.7

'''Throwaway converter of contacts from M$ CSV (Nokia) to VCF (Samsung)'''

from collections import namedtuple
from itertools import izip
import io
import sys

class Property(namedtuple('Property', 'name values')):
    def __str__(self):
        return unicode(self).encode('utf-8')

    def __unicode__(self):
        values = u';'.join(u'' if v is None else v for v in self.values)
        return u'{}:{}'.format(self.name, values)

class VcfContact(object):
    def __init__(self):
        self.properties = []

    def __str__(self):
        return unicode(self).encode('utf-8')

    def __unicode__(self):
        props = map(unicode, self.properties)
        return u'\n'.join([u'BEGIN:VCARD', u'VERSION:2.1'] + props + [u'END:VCARD'])

CSV_BUSINESS_PHONE = 'Business Phone'
CSV_CAR_PHONE = 'Car Phone'
CSV_E_MAIL_2_ADDRESS = 'E-mail 2 Address'
CSV_E_MAIL_3_ADDRESS = 'E-mail 3 Address'
CSV_E_MAIL_ADDRESS = 'E-mail Address'
CSV_FIRST_NAME = 'First Name'
CSV_HOME_PHONE = 'Home Phone'
CSV_HOME_STREET = 'Home Street'
CSV_LAST_NAME = 'Last Name'
CSV_MIDDLE_NAME = 'Middle Name'
CSV_MOBILE_PHONE = 'Mobile Phone'
CSV_TITLE = 'Title'

def iter_contacts(stream):
    '''Contacts in MS CSV format.'''
    it = iter(stream)
    header_fields = it.next().strip().split(',')
    for line in it:
        fields = line.strip().split(',')
        yield(dict((h, f) for h, f in izip(header_fields, fields) if f))

def csv_contact_to_vcf(contact):
    contact = contact.copy()
    vcf = VcfContact()

    print >>sys.stderr, '-----'
    print >>sys.stderr, u'\n'.join(u'   {}\t{}'.format(k, contact[k]) for k in sorted(contact))
    print >>sys.stderr, ''

    prop_n = None
    prop_fn = None
    prop_tel = []
    prop_email = []
    prop_adr = None

    if CSV_FIRST_NAME in contact:
        if not prop_n:
            prop_n = Property(name='N', values=[None, None, None, None, None])
        value = contact[CSV_FIRST_NAME]
        del contact[CSV_FIRST_NAME]
        prop_n.values[1] = value

    if CSV_MIDDLE_NAME in contact:
        if not prop_n:
            prop_n = Property(name='N', values=[None]*5)
        value = contact[CSV_MIDDLE_NAME]
        del contact[CSV_MIDDLE_NAME]
        prop_n.values[2] = value

    if CSV_LAST_NAME in contact:
        if not prop_n:
            prop_n = Property(name='N', values=[None]*5)
        value = contact[CSV_LAST_NAME]
        del contact[CSV_LAST_NAME]
        prop_n.values[0] = value

    if CSV_TITLE in contact:
        if not prop_n:
            prop_n = Property(name='N', values=[None]*5)
        value = contact[CSV_TITLE]
        del contact[CSV_TITLE]
        prop_n.values[3] = value

    csv_tel_cell_keys = [CSV_MOBILE_PHONE, CSV_CAR_PHONE, CSV_BUSINESS_PHONE]
    if any(tel in contact for tel in csv_tel_cell_keys):
        for tel in csv_tel_cell_keys:
            if tel in contact:
                value = contact[tel]
                del contact[tel]
                prop_tel.append(Property(name='TEL;CELL', values=[value]))

    if CSV_HOME_PHONE in contact:
        value = contact[CSV_HOME_PHONE]
        del contact[CSV_HOME_PHONE]
        prop_tel.append(Property(name='TEL;HOME', values=[value]))

    csv_email_keys = [CSV_E_MAIL_ADDRESS, CSV_E_MAIL_2_ADDRESS, CSV_E_MAIL_3_ADDRESS]
    if any(email in contact for email in csv_email_keys):
        for email in csv_email_keys:
            if email in contact:
                value = contact[email]
                del contact[email]
                prop_email.append(Property(name='EMAIL', values=[value]))

    if CSV_HOME_STREET in contact:
        value = contact[CSV_HOME_STREET]
        del contact[CSV_HOME_STREET]
        prop_adr = Property(name='ADR', values=[None]*7)
        prop_adr.values[2] = value

    if not prop_fn and prop_n:
        prop_fn = Property(name='FN',values=[u' '.join(v for v in prop_n.values[::-1] if v)])
    if not prop_fn and prop_email:
        prop_fn = Property(name='FN',values=[prop_email[0].values[0]])
    if not prop_fn and prop_tel:
        prop_fn = Property(name='FN',values=[prop_tel[0].values[0]])

    vcf.properties += [prop_fn]
    if prop_n:
        vcf.properties += [prop_n]
    vcf.properties += prop_tel + prop_email
    if prop_adr:
        vcf.properties += [prop_adr]

    if contact:
        print >>sys.stderr, 'There are still some fields left in contact'
        for k in sorted(contact):
            print >>sys.stderr, u'{} : {}'.format(k, contact[k])
        print >>sys.stderr, u'\n{}'.format(vcf)
        sys.exit(1)
    if any(not p for p in vcf.properties):
        print >>sys.stderr, 'some property is None'
        print >>sys.stderr, u'\n{}'.format(vcf)
        sys.exit(1)
    return vcf

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print 'specify M$ .csv file name'
        sys.exit(1)
    # Unicode with BOM
    # https://stackoverflow.com/questions/13590749/reading-unicode-file-data-with-bom-chars-in-python
    all_keys = set()
    with io.open(sys.argv[1], mode='r', encoding='utf-8-sig') as h:
        for c in iter_contacts(h):
            if not c:
                continue
            print csv_contact_to_vcf(c)
            print >>sys.stderr, ''
    print >>sys.stderr, 'DONE!'
