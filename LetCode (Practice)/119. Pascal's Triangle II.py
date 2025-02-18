#%% 119. Pascal's Triangle II
class Solution:
    def getRow(self, rowIndex: int) -> list[int]:
        itens_ant = [1]
        itens_atual = [1, 1]

        if rowIndex == 0:
            return itens_ant
        elif rowIndex == 1:
            return itens_atual

        for x in range(1, rowIndex):
            if len(itens_ant) < len(itens_atual):
                itens_ant = itens_atual
                itens_atual = []

                for y in range(len(itens_ant) + 1):

                    if y == 0 or y == len(itens_ant):
                        itens_atual.append(1)
                    else:
                        new_item = itens_ant[y] + itens_ant[y - 1]
                        itens_atual.append(new_item)

        return itens_atual



# %%
rowIndex = 13
teste = Solution().getRow(rowIndex)
print(teste)