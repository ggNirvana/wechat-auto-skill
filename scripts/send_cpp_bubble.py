# -*- coding: utf-8 -*-
import sys
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

message_text = """@也没错 冒泡排序来啦~

#include <iostream>
using namespace std;

void bubbleSort(int arr[], int n) {
    for (int i = 0; i < n - 1; i++) {
        for (int j = 0; j < n - i - 1; j++) {
            if (arr[j] > arr[j + 1]) {
                swap(arr[j], arr[j + 1]);
            }
        }
    }
}

int main() {
    int arr[] = {64, 34, 25, 12, 22, 11, 90};
    int n = sizeof(arr) / sizeof(arr[0]);
    
    bubbleSort(arr, n);
    
    cout << "排序后: ";
    for (int i = 0; i < n; i++) cout << arr[i] << " ";
    cout << endl;
    return 0;
}

时间复杂度 O(n^2)，空间复杂度 O(1) ~"""

try:
    from pyweixin import Navigator, Messages, GlobalConfig
except ImportError as e:
    print("请先安装 pyweixin: pip install -e /path/to/pywechat", file=sys.stderr)
    sys.exit(1)

Messages.send_messages_to_friend(
    friend="新的周末就要来啦",
    messages=[message_text],
    at_members=["也没错"],
    close_weixin=False,
)
print(f"已发送到群【新的周末就要来啦】并 @ 也没错")
