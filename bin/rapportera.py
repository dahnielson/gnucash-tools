#!/usr/bin/python -O
# coding=utf8

import sys
import codecs
import argparse
import datetime
from dateutil import tz
import sqlite3
import locale

locale.setlocale(locale.LC_ALL, '')

def c(i, s = False): 
    if i != 0 or s:
        return locale.currency(i, False, True)
    else:
        return ''

class MacroCommand:

    def __init__(self):
        self.commands = []

    def add_command(self, command):
        self.commands.append(command)
        
    def execute(self):
        for command in self.commands:
            command.execute()

class SetupReportCommand:
    
    def __init__(self, fp):
        self.fp = fp
        self.title = "Unnamed"
        self.creator = "Unnamed"

    def execute(self):
        self.fp.write('%!PS\n')
        self.fp.write('%%Title: ' + self.title + '\n')
        self.fp.write('%%Creator: ' + self.creator + '\n')
        self.fp.write('%%BoundingBox: 0 0 595 842\n')

        self.fp.write('/TimesRomanLatin << /Times-Roman findfont {} forall >> begin /Encoding ISOLatin1Encoding 256 array copy def currentdict end definefont pop\n')
        self.fp.write('/baseline { 12 } def\n')
        self.fp.write('/title { /TimesRomanLatin 14 selectfont } def\n')
        self.fp.write('/head { /TimesRomanLatin 10 selectfont } def\n')
        self.fp.write('/body { /TimesRomanLatin 10 selectfont } def\n')
        self.fp.write('/foot { /TimesRomanLatin 10 selectfont } def\n')
        self.fp.write('/tb { newpath lm tm baseline add baseline 4 div sub moveto rm tm baseline add baseline 4 div sub lineto stroke } def\n')
        self.fp.write('/bb { newpath lm tm baseline 4 div sub moveto rm tm baseline 4 div sub lineto stroke } def\n')
        self.fp.write('/fat { 1.5 setlinewidth } def /thin { .5 setlinewidth } def /normal { 1 setlinewidth } def\n')

        self.fp.write('/PG 1 def /pgno { /tempstr 10 string def PG tempstr cvs /PG PG 1 add def } def\n')
        self.fp.write('/concat { exch dup length 2 index length add string dup dup 4 2 roll copy length 4 -1 roll putinterval } bind def\n')

        self.fp.write('/textbox { /lm 48 def /tm 794 def /rm 547 def /bm 48 def lm tm moveto } def\n')
        self.fp.write('/linewidth { rm lm sub } def\n')
        self.fp.write('/newline { tm baseline sub /tm exch def lm tm moveto } def\n')
        self.fp.write('/page { textbox header } def /close { showpage } def\n')
        self.fp.write('/n { newline tm bm lt { close page } if } def\n')

        self.fp.write('/left { lm tm moveto show } def\n')
        self.fp.write('/center { dup stringwidth pop 2 div linewidth 2 div exch sub lm add tm moveto show } def\n')
        self.fp.write('/right { dup stringwidth pop rm exch sub tm moveto show } def\n')

class SetupHuvudbokCommand:

    def __init__(self, fp):
        self.fp = fp
        self.title = u"Huvudbok"
        self.company = ''
        self.date = datetime.date.today()
        self.fiscal_start = datetime.date(datetime.date.today().year, 1, 1)
        self.fiscal_end = datetime.date(datetime.date.today().year, 12, 31) 
        self.period_start = datetime.date(datetime.date.today().year, 1, 1)
        self.period_end = datetime.date(datetime.date.today().year, 12, 31) 
        self.code_start = None
        self.code_end = None
        self.last_num = 0 

    def execute(self):
        self.fp.write('/header {\n')
        self.fp.write('title (%s) center newline newline\n' % (self.title))
        self.fp.write('head (%s) left (Sida: ) pgno concat right newline\n' % (self.company))
        self.fp.write(u'head (Räkenskapsår: %s - %s) left (Utskrivet: %s) right newline\n' % (self.fiscal_start.strftime("%y%m%d"), self.fiscal_end.strftime("%y%m%d"), self.date.strftime("%y%m%d")))
        if self.code_start == None and self.code_end == None:
            self.fp.write('head (Konto: Alla) left (Senaste vernr: %s) right newline\n' % (self.last_num))
        else:
            self.fp.write('head (Konto: %s - %s) left (Senaste vernr: %s) right newline\n' % (self.code_start, self.code_end, self.last_num))
        self.fp.write('head (Period: %s - %s) left newline newline\n' % (self.period_start.strftime("%y%m%d"), self.period_end.strftime("%y%m%d")))
        self.fp.write('head (Konto) konto (Namn) namn newline\n')
        self.fp.write('head (Vernr) vernr (Datum) datum (Text) text (Debet) debet (Kredit) kredit (Saldo) saldo fat bb thin newline\n')
        self.fp.write('} def\n')

        self.fp.write('/konto { lm 0 add tm moveto show } def\n')
        self.fp.write('/namn { lm 48 add tm moveto show } def\n')
        self.fp.write('/vernr { lm 48 add tm moveto show } def\n')
        self.fp.write('/datum { lm 96 add tm moveto show } def\n')
        self.fp.write('/text { lm 144 add tm moveto show } def\n')
        self.fp.write('/transinfo { lm 144 add tm moveto show } def\n')
        self.fp.write('/debet { dup stringwidth pop rm 144 sub exch sub tm moveto show } def\n')
        self.fp.write('/kredit { dup stringwidth pop rm 72 sub exch sub tm moveto show } def\n')
        self.fp.write('/saldo { dup stringwidth pop rm 0 sub exch sub tm moveto show } def\n')

        self.fp.write('/ar { /tmp lm def /lm lm 96 add def tb /lm tmp def } def\n')

        self.fp.write('page\n')

class HuvudbokAccountCommand:

    def __init__(self, fp, code, name, account_type):
        self.fp = fp
        self.code = code
        self.name = name
        self.account_type = account_type
        self.balance_in = 0
        self.saldo_in = 0
        self.saldo_out = 0
        self.debit_total = 0
        self.credit_total = 0
        self.splits = []

    def add_split(self, num, date, comment, debit, credit, saldo):
        date = date.strftime('%y%m%d')
        self.splits.append((num, date, comment, c(debit), c(credit), c(saldo, True)))
    
    def execute(self):
        if len(self.splits) == 0 and self.account_type in ['INCOME', 'EXPENSE']:
            return
        self.fp.write(u"body (%s) konto (%s) namn n\n" % (self.code, self.name))
        if self.account_type in ['ASSET', 'CASH', 'RECEIVABLE', 'LIABILITY', 'EQUITY', 'PAYABLE']:
            self.fp.write(u"body (Ingående balans:) transinfo (%s) saldo n\n" % (c(self.balance_in, True)))
        self.fp.write(u"body (Ingående saldo:) transinfo (%s) saldo n n\n" % (c(self.saldo_in, True)))
        for s in self.splits:
            self.fp.write(u"body (%s) vernr (%s) datum (%s) text (%s) debet (%s) kredit (%s) saldo n\n" % s)
        self.fp.write(u"body (Omslutning:) transinfo (%s) debet (%s) kredit ar n\n" % (c(self.debit_total), c(self.credit_total)))
        self.fp.write(u"body (Utgående saldo:) transinfo (%s) saldo bb n n\n" % (c(self.saldo_out, True)))

class CloseHuvudbokCommand:

    def __init__(self, fp):
        self.fp = fp
        self.debit_total = 0
        self.credit_total = 0

    def execute(self):
        self.fp.write('foot (Huvudboksomslutning:) left (%s) debet (%s) kredit fat tb bb n\n' % (c(self.debit_total), c(self.credit_total)))
        self.fp.write('close\n')

class SetupDagbokCommand:

    def __init__(self, fp):
        self.fp = fp
        self.title = u"Verifikationslista"
        self.company = '' 
        self.date = datetime.date.today()
        self.fiscal_start = datetime.date(datetime.date.today().year, 1, 1)
        self.fiscal_end = datetime.date(datetime.date.today().year, 12, 31) 
        self.last_num = 0 
        self.num_start = None
        self.num_end = None

    def execute(self):
        self.fp.write('/header {\n')
        self.fp.write('title (%s) center newline newline\n' % (self.title))
        self.fp.write('head (%s) left (Sida: ) pgno concat right newline\n' % (self.company))
        self.fp.write(u'head (Räkenskapsår: %s - %s) left (Utskrivet: %s) right newline\n' % (self.fiscal_start.strftime("%y%m%d"), self.fiscal_end.strftime("%y%m%d"), self.date.strftime("%y%m%d")))
        if self.num_start == None and self.num_end == None:
            self.fp.write('head (Intervall: Alla) left (Senaste vernr: %s) right newline newline\n' % (self.last_num))
        else:
            self.fp.write('head (Intervall: %s - %s) left (Senaste vernr: %s) right newline newline\n' % (self.num_start, self.num_end, self.last_num))
        self.fp.write('head (Vernr) vernr (Datum) datum (Text) text newline\n')
        self.fp.write(u'head (Konto) konto (Benämning) namn (Debet) debet (Kredit) kredit fat bb thin newline\n')
        self.fp.write('} def\n')

        self.fp.write('/vernr { lm 0 add tm moveto show } def\n')
        self.fp.write('/datum { lm 48 add tm moveto show } def\n')
        self.fp.write('/text { lm 96 add tm moveto show } def\n') 
        self.fp.write('/konto { lm 48 add tm moveto show } def\n')
        self.fp.write('/namn { lm 96 add tm moveto show } def\n')
        self.fp.write('/debet { dup stringwidth pop rm 96 sub exch sub tm moveto show } def\n')
        self.fp.write('/kredit { dup stringwidth pop rm 0 sub exch sub tm moveto show } def\n')
        self.fp.write('/total { dup stringwidth pop rm 192 sub exch sub tm moveto show } def\n')

        self.fp.write('page\n')

class DagbokTransactionCommand:

    def __init__(self, fp, num, date, text):
        self.fp = fp
        self.num = num
        self.date = date
        self.text = text
        self.splits = []

    def add_split(self, code, name, debit, credit):
        self.splits.append((code, name, c(debit), c(credit))) 

    def execute(self):
        date = self.date.strftime('%y%m%d')
        self.fp.write(u"body (%s) vernr (%s) datum (%s) text n\n" % (self.num, date, self.text)) 
        for s in self.splits:
            self.fp.write(u"body (%s) konto (%s) namn (%s) debet (%s) kredit n\n" % s)
        self.fp.write('tb\n')

class CloseDagbokCommand:

    def __init__(self, fp):
        self.fp = fp
        self.trans_total = 0
        self.split_total = 0
        self.debit_total = 0
        self.credit_total = 0

    def execute(self):
        self.fp.write('foot (Antal verifikat:) left (%s) text fat tb n\n' % (self.trans_total))
        self.fp.write('foot (Antal transaktioner:) left (%s) text (Omslutning:) total (%s) debet (%s) kredit bb n\n' % (self.split_total, c(self.debit_total), c(self.credit_total)))
        self.fp.write('close\n')

class Dagbok:

    def __init__(self, connection, fp, args):
        self.__conn = connection
        self.fp = fp
        self.fiscal_start = datetime.datetime(args.fiscal_year, 1, 1, 0, 0, 0, tzinfo=tz.tzlocal())
        self.fiscal_end = datetime.datetime(args.fiscal_year, 12, 31, 23, 59, 59, tzinfo=tz.tzlocal()) 
        self.num_start = args.num_start
        self.num_end = args.num_end
        
    def execute_command(self, command):
        command.execute()

    def report(self):
        macro = MacroCommand()
        report = SetupReportCommand(self.fp)
        macro.add_command(report)
        preamble = SetupDagbokCommand(self.fp)
        preamble.company = self.__conn.execute(
            "SELECT string_val FROM slots WHERE name = ?", 
            ('options/Business/Company Name',)
        ).fetchone()[0]
        preamble.last_num = self.__conn.execute(
            "SELECT max(CAST(num AS INTEGER)) FROM transactions WHERE post_date >= ? AND post_date <= ?", 
            (self.fiscal_start.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"), self.fiscal_end.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"))
        ).fetchone()[0]
        preamble.fiscal_start = self.fiscal_start
        preamble.fiscal_end = self.fiscal_end
        if self.num_start == None and self.num_end == None:
            result = self.__conn.execute(
                "SELECT guid, num, post_date, description FROM transactions WHERE post_date >= ? AND post_date <= ? ORDER BY CAST(num AS INTEGER) ASC", 
                (self.fiscal_start.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"), self.fiscal_end.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"))
            )
        else:
            if self.num_start != None and self.num_end == None:
                self.num_end = preamble.last_num
            elif self.num_start == None and self.num_end != None:
                self.num_start = 1
            result = self.__conn.execute(
                "SELECT guid, num, post_date, description FROM transactions WHERE post_date >= ? AND post_date <= ? AND CAST(num AS INTEGER) >= ? AND CAST(num AS INTEGER) <= ? ORDER BY CAST(num AS INTEGER) ASC", 
                (self.fiscal_start.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"), self.fiscal_end.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"), self.num_start, self.num_end)
            )
        preamble.num_start = self.num_start
        preamble.num_end = self.num_end
        macro.add_command(preamble)
        close = CloseDagbokCommand(self.fp)
        for t in result:
            num = t[1]
            date = datetime.datetime(int(t[2][:4]), int(t[2][4:6]), int(t[2][6:8]), int(t[2][8:10]), int(t[2][10:12]), int(t[2][12:]), tzinfo=tz.tzutc()).astimezone(tz.tzlocal())
            text = t[3]
            transaction = DagbokTransactionCommand(self.fp, num, date, text)
            for code, name, value in self.__conn.execute("SELECT code, name, CAST(value_num AS REAL) / CAST(value_denom AS REAL) FROM splits, accounts WHERE splits.account_guid = accounts.guid AND tx_guid = ?", (t[0],)):
                name = name.replace(code, '').strip()
                debit = 0  
                credit = 0
                if value > 0:
                    debit = abs(value)
                    close.debit_total += debit 
                elif value < 0:
                    credit = abs(value)
                    close.credit_total += credit 
                transaction.add_split(code, name, debit, credit)
                close.split_total += 1
            macro.add_command(transaction)
            close.trans_total += 1
        macro.add_command(close)
        self.execute_command(macro)

class Huvudbok:

    def __init__(self, connection, fp, args):
        self.__conn = connection
        self.fp = fp
        self.fiscal_start = datetime.datetime(args.fiscal_year, 1, 1, 0, 0, 0, tzinfo=tz.tzlocal())
        self.fiscal_end = datetime.datetime(args.fiscal_year, 12, 31, 23, 59, 59, tzinfo=tz.tzlocal())
        if args.period_start == None:
            self.period_start = self.fiscal_start 
        else:
            self.period_start = datetime.datetime.strptime(args.period_start, '%y%m%d').replace(hour=0, minute=0, second=0, tzinfo=tz.tzlocal())
        if args.period_end == None:
            self.period_end = self.fiscal_end 
        else:
            self.period_end = datetime.datetime.strptime(args.period_end, '%y%m%d').replace(hour=23, minute=59, second=59, tzinfo=tz.tzlocal())
        self.code_start = args.code_start
        self.code_end = args.code_end

    def execute_command(self, command):
        command.execute()

    def report(self):
        macro = MacroCommand()
        report = SetupReportCommand(self.fp)
        macro.add_command(report)
        preamble = SetupHuvudbokCommand(self.fp)
        preamble.company = self.__conn.execute(
            "SELECT string_val FROM slots WHERE name = ?", 
            ('options/Business/Company Name',)
        ).fetchone()[0]
        preamble.last_num = self.__conn.execute(
            "SELECT max(CAST(num AS INTEGER)) FROM transactions WHERE post_date >= ? AND post_date <= ?", 
            (self.fiscal_start.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"), self.fiscal_end.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"))
        ).fetchone()[0]
        preamble.fiscal_start = self.fiscal_start
        preamble.fiscal_end = self.fiscal_end
        preamble.period_start = self.period_start
        preamble.period_end = self.period_end
        preamble.code_start = self.code_start
        preamble.code_end = self.code_end
        macro.add_command(preamble)
        close = CloseHuvudbokCommand(self.fp)
        # Select accounts that contain transactions
        for account_code, account_name, account_type, account_guid in self.__conn.execute("SELECT code, name, account_type, account_guid FROM accounts, splits WHERE accounts.guid = splits.account_guid AND code >= ? AND code <= ? GROUP BY code", (self.code_start, self.code_end)):
            account_name = account_name.replace(account_code, '').strip()
            account = HuvudbokAccountCommand(self.fp, account_code, account_name, account_type)
            # ingående balans: balans för kontot vid bokföringsårets början
            account.balance_in = self.__conn.execute(
                "SELECT total(CAST(value_num AS REAL) / CAST(value_denom AS REAL)) FROM transactions, splits WHERE transactions.guid = splits.tx_guid AND account_guid = ? AND post_date < ?", 
                (account_guid, self.fiscal_start.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"))
            ).fetchone()[0]
            # ingående saldo: saldo för kontot vid tidpunkten för det startdatum som valts vid utskrift av rapporten
            account.saldo_in = self.__conn.execute(
                "SELECT total(CAST(value_num AS REAL) / CAST(value_denom AS REAL)) FROM transactions, splits WHERE transactions.guid = splits.tx_guid AND account_guid = ? AND post_date < ?", 
                (account_guid, self.period_start.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"))
            ).fetchone()[0]
            account.saldo_out = account.saldo_in
            # Select transactions to the account in the period
            for num, date, comment, value in self.__conn.execute("SELECT num, post_date, description, CAST(value_num AS REAL) / CAST(value_denom AS REAL) FROM transactions, splits WHERE transactions.guid = splits.tx_guid AND account_guid = ? AND post_date >= ? AND post_date <= ? ORDER BY post_date", (account_guid, self.period_start.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"), self.period_end.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"))):
                date = datetime.datetime(int(date[:4]), int(date[4:6]), int(date[6:8]), int(date[8:10]), int(date[10:12]), int(date[12:]), tzinfo=tz.tzutc()).astimezone(tz.tzlocal())
                debit = 0
                credit = 0
                if value > 0:
                    debit = abs(value)
                    account.debit_total += debit
                    close.debit_total += debit
                elif value < 0:
                    credit = abs(value)
                    account.credit_total += credit
                    close.credit_total += credit
                account.saldo_out += value
                account.add_split(num, date, comment, debit, credit, account.saldo_out)
            # omslutning: summan av debet- och kreditkolumnerna för alla verifikationsrader inom kontot
            # utgående saldo: saldo för kontot vid tidpunkten för det slutdatum som valts vid utskrift av rapporten
            account.saldo_out = self.__conn.execute(
                "SELECT total(CAST(value_num AS REAL) / CAST(value_denom AS REAL)) FROM transactions, splits WHERE transactions.guid = splits.tx_guid AND account_guid = ? AND post_date <= ?", 
                (account_guid, self.period_end.astimezone(tz.tzutc()).strftime("%Y%m%d%H%M%S"))
            ).fetchone()[0]
            macro.add_command(account)
        macro.add_command(close) 
        self.execute_command(macro)

def output_report(args):
    conn = sqlite3.connect(args.infile)
    fp = codecs.open(args.outfile, 'w', 'latin-1')
    report = args.factory(conn, fp, args)
    report.report()
    fp.close()

def main():
    parser = argparse.ArgumentParser(description=u"Generera önskad rapport")
    subparsers = parser.add_subparsers()
    parser.add_argument('infile', help=u"GnuCash-fil att generera rapport ifrån", metavar='input.gnucash') 
    parser.add_argument('outfile', help=u"PostScript-fil att skriva rapporten till", metavar='output.ps')

    dagbok_parser = subparsers.add_parser('dagbok', help=u"skriv ut verifikationslista")
    dagbok_parser.add_argument('--fiscal-year', dest='fiscal_year', metavar='yyyy', default=datetime.date.today().year, type=int, help=u"rapportens räkenskapsår")
    dagbok_parser.add_argument('--num-start', dest='num_start', metavar='num', type=int, help=u"rapportens första vernr")
    dagbok_parser.add_argument('--num-end', dest='num_end', metavar='num', type=int, help=u"rapportens sista vernr")
    dagbok_parser.set_defaults(func=output_report, factory=Dagbok)

    huvudbok_parser = subparsers.add_parser('huvudbok', help=u"skriv ut huvudbok")
    huvudbok_parser.add_argument('--fiscal-year', dest='fiscal_year', metavar='yyyy', default=datetime.date.today().year, type=int, help=u"rapportens räkenskapsår")
    huvudbok_parser.add_argument('--code-start', dest='code_start', metavar='code', default='1000', help=u"rapportens första konto")
    huvudbok_parser.add_argument('--code-end', dest='code_end', metavar='code', default='8999', help=u"rapportens sista konto")
    huvudbok_parser.add_argument('--period-start', dest='period_start', metavar='yymmdd', help=u"rapportens period börjar")
    huvudbok_parser.add_argument('--period-end', dest='period_end', metavar='yymmdd', help=u"rapportens period slutar")
    huvudbok_parser.set_defaults(func=output_report, factory=Huvudbok)

    args = parser.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()
