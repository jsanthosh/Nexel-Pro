cmake_minimum_required(VERSION 3.24)
project(NativeSpreadsheet
    VERSION 6.3.0
    LANGUAGES CXX
    HOMEPAGE_URL "https://spreadsheet.local"
    DESCRIPTION "High-performance native spreadsheet application")

if(APPLE)
    enable_language(OBJC OBJCXX)
endif()

# C++ Standard
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
# Platform-specific compiler flags
if(MSVC)
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /W3 /O2")
    # MSVC link-time optimization
    set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} /GL")
    set(CMAKE_EXE_LINKER_FLAGS_RELEASE "${CMAKE_EXE_LINKER_FLAGS_RELEASE} /LTCG /OPT:REF /OPT:ICF")
else()
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fPIC -Wall -Wextra -O3 -fvisibility=hidden -fvisibility-inlines-hidden")
    # LTO + dead-code elimination for Release builds
    if(APPLE)
        # Clang supports thin LTO
        set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} -flto=thin")
        set(CMAKE_EXE_LINKER_FLAGS_RELEASE "${CMAKE_EXE_LINKER_FLAGS_RELEASE} -flto=thin -Wl,-dead_strip")
    else()
        # GCC uses -flto (no thin variant)
        set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} -flto")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -ffunction-sections -fdata-sections")
        set(CMAKE_EXE_LINKER_FLAGS_RELEASE "${CMAKE_EXE_LINKER_FLAGS_RELEASE} -flto -Wl,--gc-sections")
    endif()
endif()

# Force native architecture (avoid universal binary issues with Homebrew Qt)
if(APPLE)
    set(CMAKE_OSX_ARCHITECTURES "${CMAKE_SYSTEM_PROCESSOR}")
endif()

# Threading support
find_package(Threads REQUIRED)

# Qt Configuration
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

find_package(Qt6 COMPONENTS
    Core
    Gui
    Widgets
    Sql
    Network
    Concurrent
    Qml
    REQUIRED
)
find_package(Qt6CorePrivate QUIET)

find_package(SQLite3 REQUIRED)

# Core Engine Sources
set(CORE_SOURCES
    src/core/Cell.cpp
    src/core/Cell.h
    src/core/Spreadsheet.cpp
    src/core/Spreadsheet.h
    src/core/FormulaEngine.cpp
    src/core/FormulaEngine.h
    src/core/FormulaFunctions.cpp
    src/core/FormulaFunctions2.cpp
    src/core/FormulaFunctions3.cpp
    src/core/FormulaFunctions4.cpp
    src/core/FormulaFunctions5.cpp
    src/core/FormulaFunctionsBatch5.cpp
    src/core/FormulaMetadata.h
    src/core/CellRange.cpp
    src/core/CellRange.h
    src/core/ConditionalFormatting.cpp
    src/core/ConditionalFormatting.h
    src/core/UndoManager.cpp
    src/core/UndoManager.h
    src/core/DependencyGraph.cpp
    src/core/DependencyGraph.h
    src/core/NumberFormat.cpp
    src/core/NumberFormat.h
    src/core/FillSeries.cpp
    src/core/FillSeries.h
    src/core/TableStyle.h
    src/core/PivotEngine.cpp
    src/core/PivotEngine.h
    src/core/SparklineConfig.h
    src/core/NamedRange.h
    src/core/MacroEngine.cpp
    src/core/MacroEngine.h
    src/core/DocumentTheme.cpp
    src/core/DocumentTheme.h
    src/core/StringPool.cpp
    src/core/StringPool.h
    src/core/ColumnStore.cpp
    src/core/ColumnStore.h
    src/core/StyleTable.cpp
    src/core/StyleTable.h
    src/core/CellProxy.h
    src/core/FormulaAST.cpp
    src/core/FormulaAST.h
    src/core/RecalcScheduler.cpp
    src/core/RecalcScheduler.h
    src/core/FilterEngine.cpp
    src/core/FilterEngine.h
)

# Database Sources
set(DATABASE_SOURCES
    src/database/DatabaseManager.cpp
    src/database/DatabaseManager.h
    src/database/DocumentRepository.cpp
    src/database/DocumentRepository.h
)

# Services Sources
set(SERVICES_SOURCES
    src/services/ClaudeService.cpp
    src/services/ClaudeService.h
    src/services/DocumentService.cpp
    src/services/DocumentService.h
    src/services/CsvService.cpp
    src/services/CsvService.h
    src/services/XlsxService.cpp
    src/services/XlsxService.h
    src/services/AutoSaveService.cpp
    src/services/AutoSaveService.h
    src/services/IncrementalSaveService.cpp
    src/services/IncrementalSaveService.h
)

# UI Sources
set(UI_SOURCES
    src/ui/MainWindow.cpp
    src/ui/MainWindow.h
    src/ui/SpreadsheetView.cpp
    src/ui/SpreadsheetView.h
    src/ui/CellDelegate.cpp
    src/ui/CellDelegate.h
    src/ui/FormulaPopupDelegate.h
    src/ui/SpreadsheetModel.cpp
    src/ui/SpreadsheetModel.h
    src/ui/FormulaBar.cpp
    src/ui/FormulaBar.h
    src/ui/Toolbar.cpp
    src/ui/Toolbar.h
    src/ui/FormatCellsDialog.cpp
    src/ui/FormatCellsDialog.h
    src/ui/FindReplaceDialog.cpp
    src/ui/FindReplaceDialog.h
    src/ui/GoToDialog.cpp
    src/ui/GoToDialog.h
    src/ui/GoToSpecialDialog.cpp
    src/ui/GoToSpecialDialog.h
    src/ui/RemoveDuplicatesDialog.cpp
    src/ui/RemoveDuplicatesDialog.h
    src/ui/SortDialog.cpp
    src/ui/SortDialog.h
    src/ui/ConditionalFormatDialog.cpp
    src/ui/ConditionalFormatDialog.h
    src/ui/DataValidationDialog.cpp
    src/ui/DataValidationDialog.h
    src/ui/ChatPanel.cpp
    src/ui/ChatPanel.h
    src/ui/ChartWidget.cpp
    src/ui/ChartWidget.h
    src/ui/ShapeWidget.cpp
    src/ui/ShapeWidget.h
    src/ui/ChartDialog.cpp
    src/ui/ChartDialog.h
    src/ui/ShapePropertiesDialog.cpp
    src/ui/ShapePropertiesDialog.h
    src/ui/ChartPropertiesPanel.cpp
    src/ui/ChartPropertiesPanel.h
    src/ui/PivotTableDialog.cpp
    src/ui/PivotTableDialog.h
    src/ui/TemplateGallery.cpp
    src/ui/TemplateGallery.h
    src/ui/SparklineDialog.cpp
    src/ui/SparklineDialog.h
    src/ui/ImageWidget.cpp
    src/ui/ImageWidget.h
    src/ui/MacroEditorDialog.cpp
    src/ui/MacroEditorDialog.h
    src/ui/PasteSpecialDialog.cpp
    src/ui/PasteSpecialDialog.h
    src/ui/GoalSeekDialog.cpp
    src/ui/GoalSeekDialog.h
    src/ui/Theme.cpp
    src/ui/Theme.h
    src/ui/VirtualScrollBar.h
)

# Data2App ChartView (optional, macOS only — auto-detected)
set(CHARTVIEW_SOURCES "")
set(HAS_DATA2APP FALSE)
if(APPLE)
    set(CHARTVIEW_DIR "${CMAKE_CURRENT_SOURCE_DIR}/../../data2app_chartview_demo")
    set(CHARTVIEW_EXTERNAL_DIR "${CHARTVIEW_DIR}/external")
    if(EXISTS "${CHARTVIEW_DIR}/source/ChartView.mm")
        list(APPEND CMAKE_PREFIX_PATH
            "${CHARTVIEW_EXTERNAL_DIR}/boost"
            "${CHARTVIEW_EXTERNAL_DIR}/rapidjson"
            "${CHARTVIEW_EXTERNAL_DIR}/skia"
            "${CHARTVIEW_EXTERNAL_DIR}/data2pixels"
            "${CHARTVIEW_EXTERNAL_DIR}/pixels2canvas"
            "${CHARTVIEW_EXTERNAL_DIR}/data2chart"
        )
        find_package(Boost         1.82.0  CONFIG QUIET EXACT)
        find_package(RapidJSON     1.1.0   CONFIG QUIET EXACT)
        find_package(Skia          118.0.0 CONFIG QUIET EXACT)
        find_package(Data2Pixels   1.0.0   CONFIG QUIET EXACT)
        find_package(Pixels2Canvas 1.0.0   CONFIG QUIET EXACT)
        find_package(Data2Chart    1.0.0   CONFIG QUIET EXACT)

        if(Boost_FOUND AND RapidJSON_FOUND AND Skia_FOUND
           AND Data2Pixels_FOUND AND Pixels2Canvas_FOUND AND Data2Chart_FOUND)
            set(HAS_DATA2APP TRUE)
            message(STATUS "Data2App ChartView: ENABLED (Metal+Skia rendering)")
            set(CHARTVIEW_SOURCES
                src/ui/NativeChartWidget.mm
                ${CHARTVIEW_DIR}/source/ChartView.mm
                ${CHARTVIEW_DIR}/source/MetalView.mm
                ${CHARTVIEW_DIR}/source/ChartManager.cpp
                ${CHARTVIEW_DIR}/source/Chart.cpp
                ${CHARTVIEW_DIR}/source/Utils.cpp
            )
        else()
            message(STATUS "Data2App ChartView: DISABLED (missing dependencies)")
        endif()
    else()
        message(STATUS "Data2App ChartView: DISABLED (not found)")
    endif()
endif()

# Main Application
set(APP_SOURCES
    src/main.cpp
)

# Create executable
add_executable(NativeSpreadsheet
    ${APP_SOURCES}
    ${CORE_SOURCES}
    ${DATABASE_SOURCES}
    ${SERVICES_SOURCES}
    ${UI_SOURCES}
    ${CHARTVIEW_SOURCES}
    resources/resources.qrc
)

# Link libraries
target_link_libraries(NativeSpreadsheet
    Qt6::Core
    Qt6::CorePrivate
    Qt6::Gui
    Qt6::Widgets
    Qt6::Sql
    Qt6::Network
    Qt6::Concurrent
    Qt6::Qml
    SQLite::SQLite3
    Threads::Threads
)

# Include directories
target_include_directories(NativeSpreadsheet PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src/core
    ${CMAKE_CURRENT_SOURCE_DIR}/src/database
    ${CMAKE_CURRENT_SOURCE_DIR}/src/services
    ${CMAKE_CURRENT_SOURCE_DIR}/src/ui
)

# macOS specific settings
if(APPLE)
    set(MACOSX_BUNDLE_ICON_FILE AppIcon.icns)
    set(APP_ICON_MACOSX ${CMAKE_CURRENT_SOURCE_DIR}/resources/AppIcon.icns)
    set_source_files_properties(${APP_ICON_MACOSX} PROPERTIES
        MACOSX_PACKAGE_LOCATION "Resources")
    target_sources(NativeSpreadsheet PRIVATE ${APP_ICON_MACOSX})
    set_target_properties(NativeSpreadsheet PROPERTIES
        MACOSX_BUNDLE TRUE
        MACOSX_BUNDLE_GUI_IDENTIFIER "com.nexel.app"
        MACOSX_BUNDLE_BUNDLE_NAME "Nexel Pro"
        MACOSX_BUNDLE_BUNDLE_VERSION "6.3.0"
        MACOSX_BUNDLE_SHORT_VERSION_STRING "6.3"
        MACOSX_BUNDLE_ICON_FILE AppIcon.icns
        INSTALL_RPATH "@executable_path/../Frameworks"
        BUILD_RPATH_USE_ORIGIN ON
    )

    # ── Data2App ChartView (only if detected) ──
    if(HAS_DATA2APP)
        target_include_directories(NativeSpreadsheet PRIVATE
            "${CHARTVIEW_DIR}/include"
            "${CHARTVIEW_DIR}/source"
        )

        target_link_libraries(NativeSpreadsheet
            Boost::headers
            RapidJSON
            skia::skia skia::skparagraph skia::skshaper skia::skunicode skia::svg
            data2pixels::data2pixels
            pixels2canvas::pixels2canvas
            data2chart::data2chart
            "-framework Metal"
            "-framework MetalKit"
            "-framework Cocoa"
            "-framework Foundation"
            "-framework QuartzCore"
            "-framework CoreGraphics"
            "-framework CoreText"
            "-framework CoreFoundation"
            "-framework CoreServices"
            "-framework UniformTypeIdentifiers"
        )

        target_compile_definitions(NativeSpreadsheet PRIVATE
            HAS_DATA2APP
            P2C_BUILD_FOR_MACOS
            P2C_USE_SKIA_ENGINE
            P2C_USE_METAL_BACKEND
            D2C_USE_SKIA_ENGINE
        )

        # Compile Data2App .cpp and .mm files as Obj-C++
        set_source_files_properties(
            src/ui/NativeChartWidget.mm
            ${CHARTVIEW_DIR}/source/ChartView.mm
            ${CHARTVIEW_DIR}/source/MetalView.mm
            ${CHARTVIEW_DIR}/source/ChartManager.cpp
            ${CHARTVIEW_DIR}/source/Chart.cpp
            ${CHARTVIEW_DIR}/source/Utils.cpp
            PROPERTIES COMPILE_FLAGS "-x objective-c++"
        )
    endif()
endif()

# Windows specific settings
if(WIN32)
    set_target_properties(NativeSpreadsheet PROPERTIES
        WIN32_EXECUTABLE TRUE
    )
endif()

# Strip debug symbols in Release builds (saves ~3MB per platform)
if(CMAKE_BUILD_TYPE STREQUAL "Release")
    if(APPLE)
        add_custom_command(TARGET NativeSpreadsheet POST_BUILD
            COMMAND ${CMAKE_STRIP} -x "$<TARGET_FILE:NativeSpreadsheet>"
            COMMENT "Stripping debug symbols from binary")
    elseif(UNIX)
        add_custom_command(TARGET NativeSpreadsheet POST_BUILD
            COMMAND ${CMAKE_STRIP} --strip-unneeded "$<TARGET_FILE:NativeSpreadsheet>"
            COMMENT "Stripping debug symbols from binary")
    endif()
endif()

# ============================================================================
# Benchmark target (optional — not built by default)
# Build with: cmake --build build --target benchmark_20m
# ============================================================================
add_executable(benchmark_20m EXCLUDE_FROM_ALL
    benchmarks/benchmark_20m.cpp
    src/core/ColumnStore.cpp
    src/core/ColumnStore.h
    src/core/StyleTable.cpp
    src/core/StyleTable.h
    src/core/StringPool.cpp
    src/core/StringPool.h
    src/core/CellRange.cpp
    src/core/CellRange.h
    src/core/Cell.cpp
    src/core/Cell.h
    src/core/FilterEngine.cpp
    src/core/FilterEngine.h
    src/core/FormulaAST.cpp
    src/core/FormulaAST.h
    src/core/DependencyGraph.cpp
    src/core/DependencyGraph.h
)
target_link_libraries(benchmark_20m
    Qt6::Core
    Threads::Threads
)
target_include_directories(benchmark_20m PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src/core
)

# ============================================================================
# Stress Test — loads real CSV, tests sort/filter/formula for crashes
# Build with: cmake --build build --target stress_test
# ============================================================================
add_executable(stress_test EXCLUDE_FROM_ALL
    benchmarks/stress_test.cpp
    ${CORE_SOURCES}
    ${DATABASE_SOURCES}
    ${SERVICES_SOURCES}
)
target_link_libraries(stress_test
    Qt6::Core Qt6::CorePrivate Qt6::Widgets Qt6::Gui Qt6::Qml Qt6::Concurrent
    SQLite::SQLite3
    Threads::Threads
)
target_include_directories(stress_test PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
    ${SQLITE3_INCLUDE_DIR}
)

# ============================================================================
# Comprehensive Test Suite
# Build: cmake --build build --target test_all
# ============================================================================
add_executable(test_all EXCLUDE_FROM_ALL
    tests/test_all.cpp
    ${CORE_SOURCES}
    ${DATABASE_SOURCES}
    ${SERVICES_SOURCES}
)
target_link_libraries(test_all
    Qt6::Core Qt6::CorePrivate Qt6::Widgets Qt6::Gui Qt6::Qml Qt6::Concurrent
    SQLite::SQLite3
    Threads::Threads
)
target_include_directories(test_all PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
)

# ============================================================================
# Regression Test Suite
# Build: cmake --build build --target test_regression
# ============================================================================
add_executable(test_regression EXCLUDE_FROM_ALL
    tests/test_regression.cpp
    ${CORE_SOURCES}
    ${DATABASE_SOURCES}
    ${SERVICES_SOURCES}
)
target_link_libraries(test_regression
    Qt6::Core Qt6::CorePrivate Qt6::Widgets Qt6::Gui Qt6::Qml Qt6::Concurrent Qt6::Network
    SQLite::SQLite3
    Threads::Threads
)
target_include_directories(test_regression PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
)

# Install
install(TARGETS NativeSpreadsheet
    BUNDLE DESTINATION . COMPONENT Runtime
    RUNTIME DESTINATION bin COMPONENT Runtime
)
