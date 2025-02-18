#%%
# Define a classe Solution e o método reverseOnlyLetters
class Solution:
    def reverseOnlyLetters(self, s: str) -> str:

        #EXEMPLO:
        # Input: s = 'a-bC-dEf-ghIj'
        # Output: 'j-Ih-gfE-dCba'

        reverse = [x for x in s if x.isalpha()]
        reverse.reverse()

        iterador_letras = iter(reverse)
        result = [next(iterador_letras) if x.isalpha() else x for x in s]

        result = ''.join(result)
        return result

#%%
s = 'a-bC-dEf-ghIj'
# Cria uma instância da classe Solution
solucao = Solution()

# Chama o método reverseOnlyLetters na instância criada
solucao.reverseOnlyLetters(s = s)
