#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "format.h"
#include "utility.h"
#include "workbook.h"
#include "worksheet.h"

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  Chart::Init(env, exports);
  Format::Init(env, exports);
  Utility::Init(env, exports);
  Workbook::Init(env, exports);
  Worksheet::Init(env, exports);

  auto colors = Napi::Object::New(env);
  colors["BLACK_COLOR"] = Napi::Number::New(env, LXW_COLOR_BLACK);
  colors["BLUE_COLOR"] = Napi::Number::New(env, LXW_COLOR_BLUE);
  colors["BROWN_COLOR"] = Napi::Number::New(env, LXW_COLOR_BROWN);
  colors["CYAN_COLOR"] = Napi::Number::New(env, LXW_COLOR_CYAN);
  colors["GRAY_COLOR"] = Napi::Number::New(env, LXW_COLOR_GRAY);
  colors["GREEN_COLOR"] = Napi::Number::New(env, LXW_COLOR_GREEN);
  colors["LIME_COLOR"] = Napi::Number::New(env, LXW_COLOR_LIME);
  colors["MAGENTA_COLOR"] = Napi::Number::New(env, LXW_COLOR_MAGENTA);
  colors["NAVY_COLOR"] = Napi::Number::New(env, LXW_COLOR_NAVY);
  colors["ORANGE_COLOR"] = Napi::Number::New(env, LXW_COLOR_ORANGE);
  colors["PINK_COLOR"] = Napi::Number::New(env, LXW_COLOR_PINK);
  colors["PURPLE_COLOR"] = Napi::Number::New(env, LXW_COLOR_PURPLE);
  colors["RED_COLOR"] = Napi::Number::New(env, LXW_COLOR_RED);
  colors["SILVER_COLOR"] = Napi::Number::New(env, LXW_COLOR_SILVER);
  colors["WHITE_COLOR"] = Napi::Number::New(env, LXW_COLOR_WHITE);
  colors["YELLOW_COLOR"] = Napi::Number::New(env, LXW_COLOR_YELLOW);
  exports["Color"] = colors;

  return exports;
}

NODE_API_MODULE(NODE_GYP_MODULE_NAME, Init)
