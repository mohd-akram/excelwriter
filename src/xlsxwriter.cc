#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "format.h"
#include "workbook.h"
#include "worksheet.h"

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  Chart::Init(env, exports);
  Format::Init(env, exports);
  Workbook::Init(env, exports);
  Worksheet::Init(env, exports);
  auto colors = Napi::Object::New(env);
  colors.Set("BLACK_COLOR", Napi::Number::New(env, LXW_COLOR_BLACK));
  colors.Set("BLUE_COLOR", Napi::Number::New(env, LXW_COLOR_BLUE));
  colors.Set("BROWN_COLOR", Napi::Number::New(env, LXW_COLOR_BROWN));
  colors.Set("CYAN_COLOR", Napi::Number::New(env, LXW_COLOR_CYAN));
  colors.Set("GRAY_COLOR", Napi::Number::New(env, LXW_COLOR_GRAY));
  colors.Set("GREEN_COLOR", Napi::Number::New(env, LXW_COLOR_GREEN));
  colors.Set("LIME_COLOR", Napi::Number::New(env, LXW_COLOR_LIME));
  colors.Set("MAGENTA_COLOR", Napi::Number::New(env, LXW_COLOR_MAGENTA));
  colors.Set("NAVY_COLOR", Napi::Number::New(env, LXW_COLOR_NAVY));
  colors.Set("ORANGE_COLOR", Napi::Number::New(env, LXW_COLOR_ORANGE));
  colors.Set("PINK_COLOR", Napi::Number::New(env, LXW_COLOR_PINK));
  colors.Set("PURPLE_COLOR", Napi::Number::New(env, LXW_COLOR_PURPLE));
  colors.Set("RED_COLOR", Napi::Number::New(env, LXW_COLOR_RED));
  colors.Set("SILVER_COLOR", Napi::Number::New(env, LXW_COLOR_SILVER));
  colors.Set("WHITE_COLOR", Napi::Number::New(env, LXW_COLOR_WHITE));
  colors.Set("YELLOW_COLOR", Napi::Number::New(env, LXW_COLOR_YELLOW));
  exports.Set("Color", colors);
  return exports;
}

NODE_API_MODULE(NODE_GYP_MODULE_NAME, Init)
