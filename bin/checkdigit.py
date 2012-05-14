#!/usr/bin/python -O
# coding=utf8

import argparse

def digits_of(n):
    return [int(d) for d in str(n)]

def luhn_checksum(number):
    digits = digits_of(number)
    odd_digits = digits[-1::-2]
    even_digits = digits[-2::-2]
    checksum = 0
    checksum += sum(odd_digits)
    for d in even_digits:
        checksum += sum(digits_of(d*2))
    return checksum % 10
 
def is_luhn_valid(number):
    return luhn_checksum(number) == 0

def calculate_luhn(partial_number):
    check_digit = luhn_checksum(int(partial_number) * 10)
    return check_digit if check_digit == 0 else 10 - check_digit

def is_length_valid(number):
    length_digit = digits_of(number)[-2]
    return digits_of(len(str(number)))[-1] == length_digit    

def calculate_length(partial_number):
    return digits_of(len(str(partial_number)) + 2)[-1]

def validate(args):
    if args.length_digit and is_length_valid(args.number) and is_luhn_valid(args.number):
        print u"Godkänd"
    elif args.length_digit != True and is_luhn_valid(args.number):
        print u"Godkänd"
    else:
        print u"Ej godkänd"

def generate(args):
    number = args.number
    if args.length_digit:
        length_digit = calculate_length(number)
        check_digit = calculate_luhn(str(number) + str(length_digit))
        print "%s%d%d" % (number, length_digit, check_digit)
    else:
        check_digit = calculate_luhn(number)
        print "%s%d" % (number, check_digit)

def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers()
    parser.add_argument('number', type=int)

    validate_parser = subparsers.add_parser('validate', help=u"validera checksumma")
    validate_parser.add_argument('-l', action='store_true', dest='length_digit', help=u"med längdsiffra")
    validate_parser.set_defaults(func=validate)

    generate_parser = subparsers.add_parser('generate', help=u"generera referensnummer")
    generate_parser.add_argument('-l', action='store_true', dest='length_digit', help=u"med längdsiffra")
    generate_parser.set_defaults(func=generate)

    args = parser.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()
