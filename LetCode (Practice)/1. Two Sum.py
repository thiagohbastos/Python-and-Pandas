#%% 1. Two Sum
class Solution:
    def twoSum(self, nums: list[int], target: int) -> list[int]:
        result = []
        for k, v in enumerate(nums):

            if target - v in nums[k + 1:] and nums.count(target - v) <= 1:
                result.append(k)
                result.append(nums.index(target - v))
                break
            elif target - v in nums[k + 1:] and nums.count(target - v) >= 1:
                result.append(k)
                nums.pop(k)
                result.append(nums.index(target - v) + 1)
                break
            else:
                continue
        
        return result



# %%
nums = [-1,-2,-3,-4,-5]
target = -8

resultado = Solution().twoSum(nums, target)
print(resultado)