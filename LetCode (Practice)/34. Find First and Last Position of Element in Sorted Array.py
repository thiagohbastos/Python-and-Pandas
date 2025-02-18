#%% 34. Find First and Last Position of Element in Sorted Array
class Solution:
    def searchRange(self, nums: list[int], target: int) -> list[int]:
        left = 0
        right = len(nums) - 1
        start = end = None

        if target not in nums:
            return [-1, -1]

        else:
            while left <= right:
                if nums[left] == target and start == None:
                    start = left
                if nums[right] == target and end == None:
                    end = right
                if start != None and end != None:
                    break
                elif start != None and end == None:
                    right -= 1
                elif end != None and start == None:
                    left += 1
                else:
                    right -= 1
                    left += 1
            
            return [start, end]
