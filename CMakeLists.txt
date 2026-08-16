cmake_minimum_required(VERSION 3.10)
project(ark_excel C)

set(CMAKE_C_STANDARD 99)

add_library(ark_excel STATIC
    src/ark_excel.c
    src/ark_zip.c
    src/ark_reader.c
    src/ark_inflate.c
)
target_include_directories(ark_excel PUBLIC ${CMAKE_CURRENT_SOURCE_DIR})

# PC 演示程序(交叉编译时注释掉)
add_executable(demo examples/demo_report.c port/port_examples.c)
target_link_libraries(demo PRIVATE ark_excel)
