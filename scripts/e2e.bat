
@echo off
@REM set CODE_TESTS_PATH="$(pwd)/client/out/test"
@REM set CODE_TESTS_WORKSPACE="$(pwd)/client/testFixture"


set CODE_TESTS_PATH=%~dp0..\client\out\test
set CODE_TESTS_WORKSPACE=%~dp0.\vbaSample

node %CODE_TESTS_PATH%\runTest"
