#include <stdio.h>

// int add (int a, int b) {
//     return a + b;
// }

int loop(int a) {
    for (int i = 1; i < 10; i++) {
        a = a + i;
    }
    return a;
}

int main(void) {
    int c = loop(5);
    printf("the adding is %d", c);
    return 0;
}
