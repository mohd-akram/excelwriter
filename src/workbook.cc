#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "format.h"
#include "workbook.h"
#include "worksheet.h"

Napi::Object Workbook::Init(Napi::Env env, Napi::Object exports) {
  Napi::Function func = DefineClass(
      env,
      "Workbook",
      {
          InstanceMethod<&Workbook::AddChart>(
              "addChart",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Workbook::AddFormat>(
              "addFormat",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Workbook::AddWorksheet>(
              "addWorksheet",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Workbook::Close>(
              "close",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
      });
  exports.Set("Workbook", func);
  return exports;
}

Workbook::Workbook(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<Workbook>(info) {
  lxw_workbook_options options = {};
  options.output_buffer = &output_buffer;
  options.output_buffer_size = &output_buffer_size;
  workbook = workbook_new_opt(NULL, &options);
}

Workbook::~Workbook() {
  if (workbook) lxw_workbook_free(workbook);
}

Napi::Value Workbook::AddChart(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  return Chart::NewInstance(
      env,
      workbook_add_chart(workbook, info[0].As<Napi::Number>().Uint32Value()));
}

Napi::Value Workbook::AddFormat(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  return Format::NewInstance(env, workbook_add_format(workbook));
}

Napi::Value Workbook::AddWorksheet(const Napi::CallbackInfo& info) {
  auto name = info[0].As<Napi::String>();
  auto env = info.Env();
  return Worksheet::NewInstance(
      env, workbook_add_worksheet(workbook, name.Utf8Value().c_str()));
}

Napi::Value Workbook::Close(const Napi::CallbackInfo& info) {
  workbook_close(workbook);
  workbook = nullptr;
  auto env = info.Env();
  return Napi::ArrayBuffer::New(
      env, output_buffer, output_buffer_size, [](Napi::Env env, void* data) {
        free(data);
      });
}
