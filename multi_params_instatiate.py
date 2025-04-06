# from perplexity
# prompt: "is there a way in python to instantiate a new param for each value in a for loop?"

# to instantiate multiple params using for loopp and to also extract its values using for loop...

params = {}

for i in range(2):
	print(i)
	params[f'data{i+1}']=i+1

# {'data1': 1, 'data2': 2}

# interesting way to do so with class:
class Params:
    pass

my_params = Params()
for i in range(5):
    setattr(my_params, f'num_value{i}', i)


for i, char in enumerate(["A", "B", "C", "D", "E"]):
# for i, char in enumerate(["A", "B", "C", "D", "E"], start = 1):
    setattr(my_params, f'char_value{i}', char)



print(my_params.num_value0, my_params.num_value1)
print(my_params.char_value0, my_params.char_value1)



for attr, value in vars(my_params).items():
    # if attr.startswith("num_value"):
        print(f"Attribute: {attr}, Value: {value}")