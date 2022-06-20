def removingDoubleSpace(value):
    return str(value).replace('  ', ' ').replace('   ', ' ')


def amount_to_million(value):
    phone = str(value)
    words = phone.split(" ")
    message = int(phone)
    numbers = {
        "01": "One",
        "02": "Two",
        "03": "Three",
        "04": "Four",
        "05": "Five",
        "06": "Six",
        "07": "Seven",
        "08": "Eight",
        "09": "Nine",
        "1": "One",
        "2": "Two",
        "3": "Three",
        "4": "Four",
        "5": "Five",
        "6": "Six",
        "7": "Seven",
        "8": "Eight",
        "9": "Nine",
        "10": "Ten",
        "11": "Eleven",
        "12": "Twelve",
        "13": "Thirteen",
        "14": "Fourteen",
        "15": "Fifteen",
        "16": "Sixteen",
        "17": "Seventeen",
        "18": "Eighteen",
        "19": "Nineteen",
        "20": "Twenty",
        "21": "Twenty-One",
        "22": "Twenty-Two",
        "23": "Twenty-Three",
        "24": "Twenty-Four",
        "25": "Twenty-Five",
        "26": "Twenty-Six",
        "27": "Twenty-Seven",
        "28": "Twenty-Eight",
        "29": "Twenty-Nine",
        "30": "Thirty",
        "31": "Thirty-One",
        "32": "Thirty-Two",
        "33": "Thirty-Three",
        "34": "Thirty-Four",
        "35": "Thirty-Five",
        "36": "Thirty-Six",
        "37": "Thirty-Seven",
        "38": "Thirty-Eight",
        "39": "Thirty-Nine",
        "40": "Forty",
        "41": "Forty-One",
        "42": "Forty-Two",
        "43": "Forty-Three",
        "44": "Forty-Four",
        "45": "Forty-Five",
        "46": "Forty-Six",
        "47": "Forty-Seven",
        "48": "Forty-Eight",
        "49": "Forty-Nine",
        "50": "Fifty",
        "51": "Fifty-One",
        "52": "Fifty-Two",
        "53": "Fifty-Three",
        "54": "Fifty-Four",
        "55": "Fifty-Five",
        "56": "Fifty-Six",
        "57": "Fifty-Seven",
        "58": "Fifty-Eight",
        "59": "Fifty-Nine",
        "60": "Sixty",
        "61": "Sixty-One",
        "62": "Sixty-Two",
        "63": "Sixty-Three",
        "64": "Sixty-Four",
        "65": "Sixty-Five",
        "66": "Sixty-Six",
        "67": "Sixty-Seven",
        "68": "Sixty-Eight",
        "69": "Sixty-Nine",
        "70": "Seventy",
        "71": "Seventy-One",
        "72": "Seventy-Two",
        "73": "Seventy-Three",
        "74": "Seventy-Four",
        "75": "Seventy-Five",
        "76": "Seventy-Six",
        "77": "Seventy-Seven",
        "78": "Seventy-Eight",
        "79": "Seventy-Nine",
        "80": "Eighty",
        "81": "Eighty-One",
        "82": "Eighty-Two",
        "83": "Eighty-Three",
        "84": "Eighty-Four",
        "85": "Eighty-Five",
        "86": "Eighty-Six",
        "87": "Eighty-Eight",
        "88": "Eighty-Eight",
        "89": "Eighty-Nine",
        "90": "Ninety",
        "91": "Ninety-One",
        "92": "Ninety-Two",
        "93": "Ninety-Three",
        "94": "Ninety-Four",
        "95": "Ninety-Five",
        "96": "Ninety-Six",
        "97": "Ninety-Seven",
        "98": "Ninety-Eight",
        "99": "Ninety-Nine",
        }



    output = ""



    if message < 100:
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message < 1000:
        first_char = str(phone[0])
        second_char = str(phone[1])
        if second_char != "0":
            hundred = " Hundred and "
        else:
            hundred = " Hundred"
        output = numbers.get(first_char, first_char) + hundred
        hundred_char = phone[-2] + phone[-1]
        if hundred_char == "00":
            hundred_char = ""
        words = [str(hundred_char)]
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message < 10000:
        first_char = str(phone[0])
        second_char = str(phone[1])
        third_char = str(phone[2])
        output = numbers.get(first_char, first_char) + " Thousand "
        if third_char != "0":
            output += numbers.get(second_char, second_char) + " Hundred and "
        else:
            output += numbers.get(second_char, second_char) + " Hundred "
        if second_char == "0":
            output = numbers.get(first_char, first_char) + " Thousand "
        hundred_char = phone[-2] + phone[-1]
        if hundred_char == "00":
            hundred_char = ""
        words = [str(hundred_char)]
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message < 100000:
        first_char = str(phone[0]) + str(phone[1])
        second_char = str(phone[2])
        third_char = str(phone[3])

        if third_char != "0":
            hundred = " Hundred and "
        else:
            hundred = " Hundred "
        output = numbers.get(first_char, first_char) + " Thousand "
        if second_char != "0":
            output += numbers.get(second_char, second_char) + hundred
        hundred_char = phone[-2] + phone[-1]
        if hundred_char == "00":
            hundred_char = ""
        words = [str(hundred_char)]
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message < 1000000:
        first_char = str(phone[0])
        second_char = str(phone[1]) + str(phone[2])
        third_char = str(phone[3])
        fourth_char = str(phone[4])
        if fourth_char != "0":
            hundred = " Hundred and "
        else:
            hundred = " Hundred"
        output = numbers.get(first_char, first_char) + " Hundred and "
        output += numbers.get(second_char, second_char) + " Thousand "
        if third_char != "0":
            output += numbers.get(third_char, third_char) + hundred
        if second_char == "00":
            output = numbers.get(first_char, first_char) + " Hundred Thousand "
        hundred_char = phone[-2] + phone[-1]
        if hundred_char == "00":
            hundred_char = ""
        words = [str(hundred_char)]
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message < 10000000:
        first_char = str(phone[0])
        second_char = str(phone[1])
        third_char = str(phone[2]) + str(phone[3])
        fourth_char = str(phone[4])

        if fourth_char != "0":
            hundred = " Hundred and "
        else:
            hundred = " Hundred "
        output = numbers.get(first_char, first_char) + " Million "
        if second_char != "0":
            output += numbers.get(second_char, second_char) + hundred
        if third_char != "00":
            output += numbers.get(third_char, third_char) + " Thousand "
        if fourth_char != "0":
            output += numbers.get(fourth_char, fourth_char) + hundred
        hundred_char = phone[-2] + phone[-1]
        if hundred_char == "00":
            hundred_char = ""
        words = [str(hundred_char)]
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message < 100000000:
        first_char = str(phone[0]) + str(phone[1])
        second_char = str(phone[2])
        third_char = str(phone[3]) + str(phone[4])
        fourth_char = str(phone[5])
        fifth_char = str(phone[6])
        if fifth_char != "0":
            hundred = " Hundred and "
        else:
            hundred = " Hundred"
        output = numbers.get(first_char, first_char) + " Million "
        if second_char != "0":
            output += numbers.get(second_char, second_char) + hundred
        if third_char != "00":
            output += numbers.get(third_char, third_char) + " Thousand "
        if fourth_char != "0":
            output += numbers.get(fourth_char, fourth_char) + hundred
        hundred_char = phone[-2] + phone[-1]
        if hundred_char == "00":
            hundred_char = ""
        words = [str(hundred_char)]
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message < 1000000000:
        first_char = str(phone[0])
        second_char = str(phone[1]) + str(phone[2])
        third_char = str(phone[3])
        fourth_char = str(phone[4]) + str(phone[5])
        fifth_char = str(phone[6])

        if fifth_char != "0":
            hundred = " Hundred and "
        else:
            hundred = " Hundred"
        if second_char != "00":
            hundredx = " Hundred and "
        else:
            hundredx = " Hundred"
        output = numbers.get(first_char, first_char) + hundredx
        if second_char != "00":
            output += numbers.get(second_char, second_char) + " Million "
        else:
            output += " Million "
        if third_char != "0":
            output += numbers.get(third_char, third_char) + hundred
        if fourth_char != "00":
            output += numbers.get(fourth_char, fourth_char) + " Thousand "
        if fifth_char != "0":
            output += numbers.get(fifth_char, fifth_char) + hundred
        hundred_char = phone[-2] + phone[-1]
        if hundred_char == "00":
            hundred_char = ""
        words = [str(hundred_char)]
        for word in words:
            output += numbers.get(word, word)
        return removingDoubleSpace(output)

    elif message >= 1000000000:
        output = "Amount is too big to process"
        return removingDoubleSpace(output)
