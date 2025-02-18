#%% 14. Longest Common Prefix

class Solution:
    def longestCommonPrefix(self, strs: list[str]) -> str:
        common = ''

        for k, v in enumerate(strs):
            temp = ''
            qtd_letras = len(v)

            if common == '' and k == 0:
                common = v
                continue

            else:
                for x in range(qtd_letras, -1, -1):
                    
                    if v[:x] == common[:x]:
                        temp = v[:x]
                        common = temp
                        break
        
        return common




#%%
strs = ["flower","flow","flight"]
resultado = Solution().longestCommonPrefix(strs=strs)

print(resultado)
