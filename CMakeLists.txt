# CMake settings for convguid

# CMake minimum version
cmake_minimum_required(VERSION 3.5)

# project name and languages
project(refguid CXX)

##############################################################################

# refguid.exe
add_executable(refguid refguid.cpp)
target_compile_definitions(refguid PRIVATE UNICODE _UNICODE)
target_link_libraries(refguid PRIVATE shlwapi)
