#%% SOLUÇÃO 1

#Time Limit Exceeded
class Solution:
    def maxArea(self, height: list[int]) -> int:
        tam = len(height)
        result = []

        for pos_1, v in enumerate(height):

            for pos_2 in range(pos_1 + 1, tam):

                menor = min(v, height[pos_2])
                result.append(menor * (pos_2 - pos_1))
        result = max(result)
        return result



#%% SOLUÇÃO 2 - APENAS ADAPTAÇÃO COM LIST COMPREHENSION

#Time Limit Exceeded
class Solution:
    def maxArea(self, height: list[int]) -> int:
        tam = len(height)

        result = [[min(v, height[pos_2]) * (pos_2 - pos_1) 
                  for pos_2 in range(pos_1 + 1, tam)]
                  for pos_1, v in enumerate(height)]
        
        result = max([max(x) if len(x) > 0 else 0 for x in result])

        return result
    


#%% SOLUÇÃO 3 - APÓS AVALIAR REPOSTA E ESTUDAR

#A lógica por trás dessa movimentação é baseada na tentativa de maximizar a área:

#Água não pode ser armazenada acima da altura da linha mais baixa.

#Portanto, ao mover o ponteiro da linha mais baixa, 
#temos a chance de encontrar uma linha mais alta 
#e assim aumentar a altura mínima do próximo possível recipiente.

class Solution:
    def maxArea(self, height: list[int]) -> int:
        esquerda = 0
        direita = len(height) - 1
        area_max = 0

        while esquerda < direita:
            area_atual = min(height[esquerda], height[direita]) * (direita - esquerda)
            area_max = max(area_max, area_atual)

            # AQUI TA O PULO DO GATO
            if height[esquerda] > height[direita]:
                direita -= 1
            else:
                esquerda += 1

        return area_max
    


# %%
exemplo = [1,2,4,3]
solucao = Solution().maxArea(exemplo)
print(solucao)
