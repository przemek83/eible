include(FetchContent)

FetchContent_Declare(
    QuaZip
    GIT_REPOSITORY https://github.com/stachenov/quazip.git
    GIT_TAG 9d3aa3ee948c1cde5a9f873ecbc3bb229c1182ee
)

FetchContent_MakeAvailable(QuaZip)
