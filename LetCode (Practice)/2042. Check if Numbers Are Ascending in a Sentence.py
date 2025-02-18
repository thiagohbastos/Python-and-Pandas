#%% 2042. Check if Numbers Are Ascending in a Sentence
class Solution:
    def areNumbersAscending(self, s: str) -> bool:
        result = True
        nums = [int(x) for x in s.split() if x.isnumeric()]

        for x in range(len(nums)):
            
            if x < len(nums) - 1 and nums[x] >= nums[x + 1]:
                result = False
                break

        return result



#%%
s = "1 box has 3 blue 4 red 6 green and 12 yellow marbles"
teste = Solution().areNumbersAscending(s)
print(teste)