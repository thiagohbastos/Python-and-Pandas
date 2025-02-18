#%%
'''
    Example 2:

    Input: s = "LVIII"
    Output: 58
    Explanation: L = 50, V= 5, III = 3.
'''


#%%
class Solution:
    def romanToInt(self, s: str) -> int:

        num_rom = {
            'I':1,
            'V':5,
            'X':10,
            'L':50,
            'C':100,
            'D':500,
            'M':1000
        }

        adapt = {
            'IV':4,
            'IX':9,
            'XL':40,
            'XC':90,
            'CD':400,
            'CM':900
        }

        lista = list(s)

        result = remover = 0

        for k, letra in enumerate(lista):

            try:
                if f'{lista[k - 1]}{letra}' in adapt.keys() and k - 1 >= 0:
                    continue
                elif f'{letra}{lista[k + 1]}' in adapt.keys():
                    result += adapt[f'{letra}{lista[k + 1]}']

                else:
                    result += num_rom[f'{letra}']
                
            except:           
                result += num_rom[f'{letra}']

        return result


solution = Solution()

s = "MMMCDXC"

resultado = solution.romanToInt(s)
print(resultado)



# %%
