#
# Generated Makefile - do not edit!
#
# Edit the Makefile in the project folder instead (../Makefile). Each target
# has a -pre and a -post target defined where you can add customized code.
#
# This makefile implements configuration specific macros and targets.


# Environment
MKDIR=mkdir
CP=cp
GREP=grep
NM=nm
CCADMIN=CCadmin
RANLIB=ranlib
CC=gcc
CCC=g++
CXX=g++
FC=gfortran
AS=as

# Macros
CND_PLATFORM=GNU_5.4-Linux
CND_DLIB_EXT=so
CND_CONF=Debug
CND_DISTDIR=dist
CND_BUILDDIR=build

# Include project Makefile
include Makefile

# Object Directory
OBJECTDIR=${CND_BUILDDIR}/${CND_CONF}/${CND_PLATFORM}

# Object Files
OBJECTFILES= \
	${OBJECTDIR}/_ext/5c0/CKhHttpL.o \
	${OBJECTDIR}/_ext/5c0/CKhParser.o \
	${OBJECTDIR}/_ext/5c0/CKhParserELAN.o \
	${OBJECTDIR}/main.o


# C Compiler Flags
CFLAGS=

# CC Compiler Flags
CCFLAGS=-static-libstdc++
CXXFLAGS=-static-libstdc++

# Fortran Compiler Flags
FFLAGS=

# Assembler Flags
ASFLAGS=

# Link Libraries and Options
LDLIBSOPTIONS=-L/home/boost/boost_1_72_0/stage/lib

# Build Targets
.build-conf: ${BUILD_SUBPROJECTS}
	"${MAKE}"  -f nbproject/Makefile-${CND_CONF}.mk ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/khakparser

${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/khakparser: ${OBJECTFILES}
	${MKDIR} -p ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}
	${LINK.cc} -o ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/khakparser ${OBJECTFILES} ${LDLIBSOPTIONS} -lpthread -lboost_system -lboost_thread

${OBJECTDIR}/_ext/5c0/CKhHttpL.o: ../CKhHttpL.cpp 
	${MKDIR} -p ${OBJECTDIR}/_ext/5c0
	${RM} "$@.d"
	$(COMPILE.cc) -g -I/home/boost/boost_1_72_0 -std=c++11 -MMD -MP -MF "$@.d" -o ${OBJECTDIR}/_ext/5c0/CKhHttpL.o ../CKhHttpL.cpp

${OBJECTDIR}/_ext/5c0/CKhParser.o: ../CKhParser.cpp 
	${MKDIR} -p ${OBJECTDIR}/_ext/5c0
	${RM} "$@.d"
	$(COMPILE.cc) -g -I/home/boost/boost_1_72_0 -std=c++11 -MMD -MP -MF "$@.d" -o ${OBJECTDIR}/_ext/5c0/CKhParser.o ../CKhParser.cpp

${OBJECTDIR}/_ext/5c0/CKhParserELAN.o: ../CKhParserELAN.cpp 
	${MKDIR} -p ${OBJECTDIR}/_ext/5c0
	${RM} "$@.d"
	$(COMPILE.cc) -g -I/home/boost/boost_1_72_0 -std=c++11 -MMD -MP -MF "$@.d" -o ${OBJECTDIR}/_ext/5c0/CKhParserELAN.o ../CKhParserELAN.cpp

${OBJECTDIR}/main.o: main.cpp 
	${MKDIR} -p ${OBJECTDIR}
	${RM} "$@.d"
	$(COMPILE.cc) -g -I/home/boost/boost_1_72_0 -std=c++11 -MMD -MP -MF "$@.d" -o ${OBJECTDIR}/main.o main.cpp

# Subprojects
.build-subprojects:

# Clean Targets
.clean-conf: ${CLEAN_SUBPROJECTS}
	${RM} -r ${CND_BUILDDIR}/${CND_CONF}
	${RM} ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/khakparser

# Subprojects
.clean-subprojects:

# Enable dependency checking
.dep.inc: .depcheck-impl

include .dep.inc
