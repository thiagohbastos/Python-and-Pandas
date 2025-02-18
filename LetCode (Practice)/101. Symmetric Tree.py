#%% 101. Symmetric Tree

# Definition for a binary tree node.
# class TreeNode:
#     def __init__(self, val=0, left=None, right=None):
#         self.val = val
#         self.left = left
#         self.right = right

class Solution:
    def isSymmetric(self, root: list[int]) -> bool:
        result = True

        if len(root) % 2 == 0:
            return False

        else:
            #print(len(root))
            pos_atual = 1
            tam_atual = 2

            while pos_atual + tam_atual <= len(root):
                #print(root[pos_atual : pos_atual + tam_atual])
                #print(root[pos_atual + tam_atual - 1: pos_atual - 1: - 1])

                if root[pos_atual : pos_atual + tam_atual] == root[pos_atual + tam_atual - 1: pos_atual - 1: - 1]:
                    pos_atual += tam_atual
                    tam_atual *= 2
                    continue
                else:
                    return False

        return result


# %%
root = [1,2,2,3,4,4,3,8,5,7,33,33,7,5,8]
teste = Solution()
teste = teste.isSymmetric(root)
print(teste)
