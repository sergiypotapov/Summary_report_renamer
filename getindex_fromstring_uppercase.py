import string


def get_index_from_letter(letter):

    alphabet = string.ascii_uppercase

    index = alphabet.index(letter)

    return (index)

def get_letter_from_index(index):
    letter = string.ascii_uppercase[index]

    return (letter)