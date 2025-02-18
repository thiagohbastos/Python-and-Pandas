#%% 
class Solution:
    def has_duplicates(self, lst): 
        return len([x for x in set(lst) if lst.count(x) > 1]) > 0
    
    def maximumSubarraySum(self, nums: list[int], k: int) -> int:
        start = 0
        current_sum = 0
        max_sum = 0

        for end in range(k, len(nums) + 1):
            if self.has_duplicates(nums[start:end]):
                start += 1
                continue
            current_sum = sum(nums[start:end])
            if current_sum > max_sum :
                max_sum = max(max_sum,current_sum)
                #print(nums[start:end], current_sum)
            start += 1
        
        return max_sum



#%%
nums = [1,5,4,2,9,9,9]
k = 3

resultado = Solution().maximumSubarraySum(nums=nums, k=k)
print(resultado)

# %%
