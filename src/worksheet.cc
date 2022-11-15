#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "format.h"
#include "worksheet.h"

Napi::Object Worksheet::Init(Napi::Env env, Napi::Object exports) {
  auto func = DefineClass(
      env,
      "Worksheet",
      {
          InstanceMethod<&Worksheet::InsertChart>(
              "insertChart",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Worksheet::InsertImage>(
              "insertImage",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Worksheet::SetColumn>(
              "setColumn",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Worksheet::WriteDatetime>(
              "writeDatetime",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Worksheet::WriteNumber>(
              "writeNumber",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
          InstanceMethod<&Worksheet::WriteString>(
              "writeString",
              static_cast<napi_property_attributes>(napi_writable |
                                                    napi_configurable)),
      });

  auto data = env.GetInstanceData<Napi::ObjectReference>();

  if (!data) {
    data = new Napi::ObjectReference();
    *data = Napi::Persistent(Napi::Object::New(env));
    env.SetInstanceData(data);
  }

  data->Set("WorksheetConstructor", func);

  return exports;
}

Worksheet::Worksheet(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<Worksheet>(info) {
  worksheet = info[0].As<Napi::External<lxw_worksheet>>().Data();
}

Napi::Value Worksheet::NewInstance(Napi::Env env, lxw_worksheet* worksheet) {
  return env.GetInstanceData<Napi::ObjectReference>()
      ->Get("WorksheetConstructor")
      .As<Napi::Function>()
      .New({Napi::External<lxw_worksheet>::New(env, worksheet)});
}

Napi::Value Worksheet::InsertChart(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_insert_chart(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value(),
                         Chart::Get(info[2]));
  return env.Undefined();
}

Napi::Value Worksheet::InsertImage(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  auto buffer = info[2].As<Napi::Uint8Array>();
  worksheet_insert_image_buffer(worksheet,
                                info[0].As<Napi::Number>(),
                                info[1].As<Napi::Number>().Uint32Value(),
                                buffer.Data(),
                                buffer.ByteLength());
  return env.Undefined();
}

Napi::Value Worksheet::SetColumn(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_column(worksheet,
                       info[0].As<Napi::Number>().Uint32Value(),
                       info[1].As<Napi::Number>().Uint32Value(),
                       info[2].As<Napi::Number>(),
                       Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteDatetime(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  auto date = info[2].As<Napi::Object>();
  auto offset = date.Get("getTimezoneOffset")
                    .As<Napi::Function>()
                    .Call(date, {})
                    .As<Napi::Number>()
                    .Int32Value() *
                60;
  worksheet_write_unixtime(worksheet,
                           info[0].As<Napi::Number>(),
                           info[1].As<Napi::Number>().Uint32Value(),
                           info[2].As<Napi::Date>() / 1000 - offset,
                           Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteNumber(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_number(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value(),
                         info[2].As<Napi::Number>(),
                         Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteString(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_string(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value(),
                         info[2].As<Napi::String>().Utf8Value().c_str(),
                         Format::Get(info[3]));
  return env.Undefined();
}
