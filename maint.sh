#!/bin/sh

# 2021-03-22 Jerker Bäck

find . -type f \( -name "*.h" -o -name "*.cpp" -o -name "*.c" \) -exec dos2unix -k -q {} \; # .htm *.html
find . -type f -name "resource.h" -exec unix2dos -k -q {} \;
# find . -type f \( -name "*.h" -o -name "*.cpp" -o -name "*.c" \) -exec clang-format -i -style=file {} \;

# 'file' can be used to determine the character encoding
# file *

# find . -type f \( -name "*.h" -o -name "*.cpp" -o -name "*.inl" \) -exec sed -i 's/\boldtext\b/newtext/g' {} \;
# sed -i "s/\<oldtext\>/newtext/g"