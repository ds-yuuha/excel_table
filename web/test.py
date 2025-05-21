num = 160
int_value = int('{:.0f}'.format(num))
float_value = num - int_value
float_value = float('{:.2f}'.format(float_value))
sexa_value = round(float_value*10/6,2)
decimal_value = round(float_value*0.6,2)

total_value = int_value + sexa_value
total_value = '{:.2f}'.format(total_value)
print(total_value)
total_value = int_value + decimal_value
total_value = '{:.2f}'.format(total_value)
print(total_value)