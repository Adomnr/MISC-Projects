def is_integer(variable):
    return isinstance(variable, int)

# Test the function
print(is_integer(42))  # True
print(is_integer("42"))  # False
print(is_integer(3.14))  # False