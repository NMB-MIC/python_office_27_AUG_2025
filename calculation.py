def plus(input_1 , input_2):
    """ plus 2 digit"""
    result = input_1 + input_2 
    if result < 10:
        msg = "ok"
    else:
        msg = "ng"
    return result, msg

def minus(input_1 , input_2):
    """ minus 2 digit"""
    result = input_1 - input_2 
    if result < 10:
        msg = "ok"
    else:
        msg = "ng"
    return result, msg