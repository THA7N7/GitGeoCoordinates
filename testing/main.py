# Product 1000g
#
# Ingredient 10.5g

def convert_input(m):
    try:
        m = float(m)
    except:
        m = 0
    return m


def calc_ingredient(m):
    # 1000 : 10.5 = m : x
    return 10.5 * m / 1000


def main():
    m = convert_input(input())
    print('Инградиент: {}'.format(calc_ingredient(m)))


if __name__ == '__main__':
    main()
