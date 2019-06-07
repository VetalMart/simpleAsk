def create_range(pos1, pos2, raw):
    """
    Превращает координаты ячеек екселя, в список имен ячеек.
    Нужна для указания места, куда ложить данные. 
    Принимает две буквенные координаты, и номер строки.
    pos1 - первая координата строки
    pos2 - вторая координата строки, может состоять из 2 букв
    raw - номер строки
    """

    if len(pos2) == 1:
        #если одна буква  
        ex_range = ['{0}{1}'.format(chr(i), raw) for i in range(
                        ord(pos1), ord(pos2)+1)]
        #если 2 буквы
    else:
        ex_range = ['{0}{1}'.format(chr(i), raw) for i in range(
                            ord(pos2[0]), 91)]
        double_range = ['{0}{1}{2}'.format('A', chr(i), raw) for i in range(
                            65, ord(pos2[1])+1)]
        ex_range.extend(double_range)

    return ex_range

print(create_range('A','A   I', 10))