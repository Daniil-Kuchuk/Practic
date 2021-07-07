from Reports import Reports


def main():
    p = Reports(path='D:\\test')
    # prefixes = input('Введите префиксы: ')
    prefixes = 'Pn - давление, V - скорость, Gh - обозначение'
    new_prefix = [prefix.split('-') for prefix in prefixes.split(',')]
    pref = {item[0].strip(): item[1].strip() for item in new_prefix}

    p.add_prefix(pref)
    p.create_slide()
main()