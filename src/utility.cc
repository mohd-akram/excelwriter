#include <napi.h>
#include <xlsxwriter.h>

#include "utility.h"

Napi::Object Utility::Init(Napi::Env env, Napi::Object exports) {
  exports.Set("cell", Napi::Function::New(env, Cell));
  return exports;
}

Napi::Array Utility::Cell(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  auto cell = info[0].As<Napi::String>().Utf8Value();
  auto rowcol = Napi::Array::New(env, 2);
  rowcol[0u] = lxw_name_to_row(cell.c_str());
  rowcol[1u] = lxw_name_to_col(cell.c_str());
  return rowcol;
}
