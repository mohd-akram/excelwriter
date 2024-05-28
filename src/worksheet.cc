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
          InstanceMethod<&Worksheet::InsertChart>("insertChart",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::InsertImage>("insertImage",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::MergeRange>("mergeRange",
                                                 napi_default_method),
          InstanceMethod<&Worksheet::SetColumn>("setColumn",
                                                napi_default_method),
          InstanceMethod<&Worksheet::SetRow>("setRow", napi_default_method),
          InstanceMethod<&Worksheet::SetFooter>("setFooter",
                                                napi_default_method),
          InstanceMethod<&Worksheet::SetHeader>("setHeader",
                                                napi_default_method),
          InstanceMethod<&Worksheet::SetSelection>("setSelection",
                                                   napi_default_method),
          InstanceMethod<&Worksheet::WriteBoolean>("writeBoolean",
                                                   napi_default_method),
          InstanceMethod<&Worksheet::WriteDatetime>("writeDatetime",
                                                    napi_default_method),
          InstanceMethod<&Worksheet::WriteFormula>("writeFormula",
                                                   napi_default_method),
          InstanceMethod<&Worksheet::WriteNumber>("writeNumber",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::WriteString>("writeString",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::WriteURL>("writeURL", napi_default_method),
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

Napi::Value Worksheet::New(Napi::Env env, lxw_worksheet* worksheet) {
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

Napi::Value Worksheet::MergeRange(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_merge_range(worksheet,
                        info[0].As<Napi::Number>(),
                        info[1].As<Napi::Number>().Uint32Value(),
                        info[2].As<Napi::Number>(),
                        info[3].As<Napi::Number>().Uint32Value(),
                        info[4].As<Napi::String>().Utf8Value().c_str(),
                        Format::Get(info[5]));
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

Napi::Value Worksheet::SetRow(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_row(worksheet,
                    info[0].As<Napi::Number>(),
                    info[1].As<Napi::Number>(),
                    Format::Get(info[2]));
  return env.Undefined();
}

Napi::Value Worksheet::SetFooter(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_footer(worksheet,
                       info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Worksheet::SetHeader(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_header(worksheet,
                       info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Worksheet::SetSelection(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_selection(worksheet,
                          info[0].As<Napi::Number>(),
                          info[1].As<Napi::Number>().Uint32Value(),
                          info[2].As<Napi::Number>(),
                          info[3].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Worksheet::WriteBoolean(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_boolean(worksheet,
                          info[0].As<Napi::Number>(),
                          info[1].As<Napi::Number>().Uint32Value(),
                          info[2].As<Napi::Boolean>(),
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

Napi::Value Worksheet::WriteFormula(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_formula(worksheet,
                          info[0].As<Napi::Number>(),
                          info[1].As<Napi::Number>().Uint32Value(),
                          info[2].As<Napi::String>().Utf8Value().c_str(),
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

Napi::Value Worksheet::WriteURL(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_url(worksheet,
                      info[0].As<Napi::Number>(),
                      info[1].As<Napi::Number>().Uint32Value(),
                      info[2].As<Napi::String>().Utf8Value().c_str(),
                      Format::Get(info[3]));
  return env.Undefined();
}
