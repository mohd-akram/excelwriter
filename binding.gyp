{
  'targets': [
    {
      'include_dirs': [
        '<!(node -p "require(\'node-addon-api\').include_dir")',
        'src/libxlsxwriter/include'
      ],
      'cflags!': [ '-fno-exceptions' ],
      'cflags_cc!': [ '-fno-exceptions' ],
      'conditions': [
        ['OS == "win"', {
          'defines': [
            '_HAS_EXCEPTIONS=1'
          ],
          'msvs_settings': {
            'VCCLCompilerTool': {
              'ExceptionHandling': 1
            },
          },
        }],
        ['OS == "mac"', {
          'xcode_settings': {
            'GCC_ENABLE_CPP_EXCEPTIONS': 'YES',
            'CLANG_CXX_LIBRARY': 'libc++',
            'MACOSX_DEPLOYMENT_TARGET': '10.7',
          },
        }],
      ],
      'target_name': 'xlsxwriter',
      'dependencies': [ 'libxlsxwriter' ],
      'sources': [
        'src/chart.cc',
        'src/format.cc',
        'src/workbook.cc',
        'src/worksheet.cc',
        'src/xlsxwriter.cc',
      ]
    },
    {
      'target_name': 'libxlsxwriter',
      'type': 'static_library',
      'defines': [ 'USE_STANDARD_TMPFILE' ],
      'conditions': [
        ['OS != "win"', {
          'defines': [ 'USE_FMEMOPEN' ],
        }],
      ],
      'dependencies': [ 'minizip', 'md5' ],
      'include_dirs': [ 'src/libxlsxwriter/include' ],
      'sources': [
        'src/libxlsxwriter/src/app.c',
        'src/libxlsxwriter/src/chart.c',
        'src/libxlsxwriter/src/chartsheet.c',
        'src/libxlsxwriter/src/comment.c',
        'src/libxlsxwriter/src/content_types.c',
        'src/libxlsxwriter/src/core.c',
        'src/libxlsxwriter/src/custom.c',
        'src/libxlsxwriter/src/drawing.c',
        'src/libxlsxwriter/src/format.c',
        'src/libxlsxwriter/src/hash_table.c',
        'src/libxlsxwriter/src/metadata.c',
        'src/libxlsxwriter/src/packager.c',
        'src/libxlsxwriter/src/relationships.c',
        'src/libxlsxwriter/src/shared_strings.c',
        'src/libxlsxwriter/src/styles.c',
        'src/libxlsxwriter/src/table.c',
        'src/libxlsxwriter/src/theme.c',
        'src/libxlsxwriter/src/utility.c',
        'src/libxlsxwriter/src/vml.c',
        'src/libxlsxwriter/src/workbook.c',
        'src/libxlsxwriter/src/worksheet.c',
        'src/libxlsxwriter/src/xmlwriter.c',
      ]
    },
    {
      'target_name': 'minizip',
      'type': 'static_library',
      'conditions': [
        ['OS == "win"', {
          'sources': [ 'src/libxlsxwriter/third_party/minizip/iowin32.c' ],
        }],
      ],
      'sources': [
        'src/libxlsxwriter/third_party/minizip/ioapi.c',
        'src/libxlsxwriter/third_party/minizip/zip.c',
      ]
    },
    {
      'target_name': 'md5',
      'type': 'static_library',
      'sources': [ 'src/libxlsxwriter/third_party/md5/md5.c' ]
    }
  ]
}
