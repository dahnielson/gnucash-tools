#!/usr/bin/python -O

import argparse

def luhn_checksum(number):
    def digits_of(n):
        return [int(d) for d in str(n)]
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

def output(args):
    check_digit = calculate_luhn(args.number)
    print "%s%d" % (args.number, check_digit)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('number')
    parser.set_defaults(func=output)
    args = parser.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()
