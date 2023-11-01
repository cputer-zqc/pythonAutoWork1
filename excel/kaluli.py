def z2k(val, weight):
    '''
    脂肪转卡路里
    :param val:
    :param weight:
    :return:
    '''
    baseG = 100 # 100g
    return val * 9 * weight / baseG

def t2k(val, weight):
    '''
    碳水化合物转卡路里
    :param val:
    :param weight:
    :return:
    '''
    baseG = 100 # 100g
    return val * 4 * weight / baseG

def d2k(val, weight):
    '''
    碳水化合物转卡路里
    :param val:
    :param weight:
    :return:
    '''
    baseG = 100 # 100g
    return val * 4 * weight / baseG

def n2k( val, weight ):
    '''
    千焦能量转为千卡
    :param val:
    :param weight:
    :return:
    '''
    baseG = 100 # 100g
    return val * 0.2389 *  weight/ baseG

if __name__ == "__main__":
    nl = 1660.0 # 能量 Kj
    dbz = 8.5 # 蛋白质 g
    zf = 0 # 脂肪 g
    tshhw = 89.2 # 碳水化合物 g
    weight = 10.0
    kaluli = n2k(nl, weight) + d2k(dbz, weight) + z2k(zf, weight) + t2k(tshhw, weight)
    print("能量值:" + str(kaluli *2))

