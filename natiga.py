#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Saves students (by seats and by school) results and ranks from http://natiga.nezakr.org

# Usage examples:
#   - ./natiga.py -f html --seats {20001..20023}
#   - ./natiga.py -f excel sqlite --school <link-to-school>

# FIXME: When saving students of different divisions, seperate subjects when listing.

import bs4, argparse, sqlite3, json, sys, urllib3, openpyxl

import urllib.parse as urlparse
from urllib.parse import urlencode

http = None
students = []
total = 0
num_students = 0

class School:
    def __init__(self, link):
        try:
            self.link = link
            page = bs4.BeautifulSoup(open_link('POST', self.link, page=1, order='mark_desc'), 'lxml')
            global total
            total += int(page.find('h4').text.strip().split(' ')[0])
            res = page.find('tbody')
            if res == None: raise ValueError()
            # Page numbers
            num = page.find(attrs={'class': 'pagination'})
            #current = int(num.find(attrs={'class': 'active'}).text)
            last = int(num.findAll('li')[-1].text)
            # Registering students
            for page in range(2, last+1):
                students = res.findAll('tr')
                for student in students:
                    try: seat = int(student.findAll('td')[1].a['href'].split('=')[-1])
                    except IndexError: continue # Just a harmless dummy tr
                    Student(seat)
                res = bs4.BeautifulSoup(open_link('POST', self.link, page=page, order='mark_desc'), 'lxml').find('tbody')
        except ValueError:
            print('Invalid School Link')

class Student:
    def __init__(self, seat, from_school=False):
        try:
            link = 'natiga.nezakr.org/index.php?t=num&k=%d' % seat
            page = bs4.BeautifulSoup(open_link('GET', link), 'lxml')
            res = page.findAll('tbody')
            if res == None: raise ValueError()
            # Student data
            self.info = {}
            self.success = True
            data = res[0].findAll('td')
            i = 0
            while i < 12:
                self.info[data[i*2].text.strip()] = data[i*2+1].text.strip()
                if i == 2: i += 2
                if i == 7:
                    if data[15].text.strip() == 'ناجح':
                        i += 3
                    else:
                        self.success = False
                        i += 1
                if not self.success and i == 9: break
                i += 1
            # Ensure not a duplicate
            if not from_school and options.schools:
                for student in students:
                    if student.info['رقم الجلوس'] == self.info['رقم الجلوس']: return None
            # Student marks
            self.marks = {}
            data = res[1].findAll('td')
            for i in range(0, len(data)//3):
                if not self.success and not data[i*3+1].text.strip(): continue # no second try
                self.marks[data[i*3].text.strip()] = data[i*3+1].text.strip()
            # Student ranks
            if self.success:
                data = res[2].findAll('td')
                self.ranks = {
                    #'الترتيب على الجمهورية': data[2].text.strip(),
                    'الترتيب على الشعبة': data[5].text.strip(),
                    #'الترتيب على المحافظة': data[8].text.strip()
                }
            else: self.ranks = {'الترتيب على الشعبة': '0'}
            p()
            students.append(self)
        except ValueError:
            print('Invalid Seat Number: %d' % seat if seat else 'Invalid Student Link: %s' % link)
        except AssertionError: pass

def parse_args():
    parser = argparse.ArgumentParser(description="Ranks students' results", epilog='(C) 2018 -- Amr Ayman')

    parser.add_argument('--seats', nargs='+', type=int, help='Student seat numbers')
    parser.add_argument('--schools', nargs='+', help='Link to a school')
    parser.add_argument('-o', '--outfile', required=True, help='Output filename')
    parser.add_argument('-f', default=['excel'], nargs='+', choices=['html', 'excel', 'sqlite'],
        help='Output file format. You can specify multiple, e.g: -f html excel ..', dest='fileformats')
    options = parser.parse_args()
    # Options stuff
    global total
    if options.seats:
        options.seats = set(options.seats)
        total += len(options.seats)
    if options.schools:
        options.schools = set(options.schools)
    if not (options.schools or options.seats):
        parser.error('No data given, add schools or seats')
    options.fileformats = set(options.fileformats)
    return options

def open_link(method, link, **params):
    if params and method in ('POST', 'PUT'):
        parts = list(urlparse.urlparse(link))
        query = dict(urlparse.parse_qsl(parts[4]))
        query.update(params)
        parts[4] = urlencode(query)
        link = urlparse.urlunparse(parts)
    try:
        if method in ('POST', 'PUT'):
            return http.request(method, link, redirect=False).data
        else:
            return http.request(method, link, redirect=False, **params).data
    except Exception as e:
        print('Link cannot be opened: %s' % e, file=sys.stderr)
        raise AssertionError()

def p():
    global num_students
    num_students += 1
    msg = 'Collected %d out of %s student(s) ..' % (num_students, total)
    sys.stdout.write(msg + chr(8) * len(msg))
    sys.stdout.flush()

if __name__ == '__main__':
    try:
        # Initializing environment
        http = urllib3.PoolManager(retries=4)
        options = parse_args()
        # Collecting data
        if options.schools:
            for school in options.schools:
                School(school)
        if options.seats:
            for seat in options.seats:
                Student(seat=seat)
        # Sorting according to marks. If one school only, sorted already.
        # Also, have sorts for seperate divisions.
        sorts = []
        divs = ['all']
        if options.seats or len(options.schools) > 1:
            students.sort(key=lambda student: float(student.info['المجموع']), reverse=True)
        sorts.append(students)
        for division in ['علمي رياضيات', 'علمي علوم', 'أدبي']:
            tmp = [student for student in students if list(student.info.values())[-1] == division ]
            if len(tmp) == 0: continue
            divs.append(division)
            d = []
            for subject in tmp[0].marks:
                if tmp[0].marks[subject] == 'غير مقرر': d.append(subject)
            # To avoid a runtime error
            for subject in d:
                for student in tmp: del student.marks[subject]
            tmp.sort(key=lambda student: float(student.info['المجموع']), reverse=True)
            sorts.append(tmp)
        # Writing data
        for d, sort in enumerate(sorts):
            headers = []
            name = '%s-%s' % (options.outfile, divs[d])
            for h in sort[0].info.keys(), sort[0].marks.keys(), sort[0].ranks.keys():
                headers.extend(h)
            for format in options.fileformats:
                if format == 'html':
                    file = name + '.html'
                    with open(file, 'w') as f:
                        f.write('<table><tr>')
                        for header in headers:
                            f.write('<th>%s</th>' % header)
                        f.write('</tr>')
                        for student in sort:
                            f.write('<tr>')
                            for h in student.info.values(), student.marks.values(), student.ranks.values():
                                for v in h: f.write('<td>%s</td>' % v)
                            f.write('</tr>')
                        f.write('</table>')
                elif format == 'excel':
                    file = name + '.xsl'
                    wb = openpyxl.Workbook()
                    wsl = wb.active
                    for i, header in enumerate(headers, 1):
                        wsl.cell(row=1, column=i, value=header)
                    for x, student in enumerate(sort, 2):
                        data = []
                        for h in student.info.values(), student.marks.values(), student.ranks.values():
                            data.extend(h)
                        for i, datum in enumerate(data, 1):
                            wsl.cell(row=x, column=i, value=datum)
                    wb.save(file)
                elif format == 'sqlite':
                    file = options.outfile + '.db'
                    conn = sqlite3.connect(file)
                    c = conn.cursor()
                    tmp = ['"%s" string' % header for header in headers]
                    c.execute('create table results_%s (%s)' % (divs[d], ', '.join(tmp)))
                    for student in sort:
                        tmp = []
                        for h in student.info.values(), student.marks.values(), student.ranks.values():
                            tmp.extend(h)
                        c.execute('insert into results values (%s)' % ','.join('?'*len(tmp)), tmp)
                    conn.commit()
                    c.close()
                    conn.close()
                print('Written to %s!' % file)
    except KeyboardInterrupt:
        print('\nInterrupted. Exiting ..', file=sys.stderr)
        sys.exit(1)
