#include <napi.h>
#include <xlsxwriter.h>

#include "format.h"

Napi::Object Format::Init(Napi::Env env, Napi::Object exports) {
  auto func = DefineClass(
      env,
      "Format",
      {InstanceMethod<&Format::SetBgColor>(
           "setBgColor",
           static_cast<napi_property_attributes>(napi_writable |
                                                 napi_configurable)),
       InstanceMethod<&Format::SetFgColor>(
           "setFgColor",
           static_cast<napi_property_attributes>(napi_writable |
                                                 napi_configurable)),
       InstanceMethod<&Format::SetBorderColor>(
           "setBorderColor",
           static_cast<napi_property_attributes>(napi_writable |
                                                 napi_configurable)),
       InstanceMethod<&Format::SetFontColor>(
           "setFontColor",
           static_cast<napi_property_attributes>(napi_writable |
                                                 napi_configurable)),
       InstanceMethod<&Format::SetBold>("setBold",
                                        static_cast<napi_property_attributes>(
                                            napi_writable | napi_configurable)),
       InstanceMethod<&Format::SetBorder>(
           "setBorder",
           static_cast<napi_property_attributes>(napi_writable |
                                                 napi_configurable)),
       InstanceMethod<&Format::SetNumFormat>(
           "setNumFormat",
           static_cast<napi_property_attributes>(napi_writable |
                                                 napi_configurable)),

       StaticValue("NONE_BORDER",
                   Napi::Number::New(env, LXW_BORDER_NONE),
                   napi_default),
       StaticValue("THIN_BORDER",
                   Napi::Number::New(env, LXW_BORDER_THIN),
                   napi_default),
       StaticValue("MEDIUM_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM),
                   napi_default),
       StaticValue("DASHED_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DASHED),
                   napi_default),
       StaticValue("DOTTED_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DOTTED),
                   napi_default),
       StaticValue("THICK_BORDER",
                   Napi::Number::New(env, LXW_BORDER_THICK),
                   napi_default),
       StaticValue("DOUBLE_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DOUBLE),
                   napi_default),
       StaticValue("HAIR_BORDER",
                   Napi::Number::New(env, LXW_BORDER_HAIR),
                   napi_default),
       StaticValue("MEDIUM_DASHED_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM_DASHED),
                   napi_default),
       StaticValue("DASH_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DASH_DOT),
                   napi_default),
       StaticValue("MEDIUM_DASH_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM_DASH_DOT),
                   napi_default),
       StaticValue("DASH_DOT_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DASH_DOT_DOT),
                   napi_default),
       StaticValue("MEDIUM_DASH_DOT_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM_DASH_DOT_DOT),
                   napi_default),
       StaticValue("SLANT_DASH_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_SLANT_DASH_DOT),
                   napi_default)});

  auto data = env.GetInstanceData<Napi::ObjectReference>();

  if (!data) {
    data = new Napi::ObjectReference();
    *data = Napi::Persistent(Napi::Object::New(env));
    env.SetInstanceData(data);
  }

  data->Set("FormatConstructor", func);
  exports.Set("Format", func);

  return exports;
}

Format::Format(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<Format>(info) {
  format = info[0].As<Napi::External<lxw_format>>().Data();
}

lxw_format* Format::Get(Napi::Value value) {
  return value.IsUndefined() || value.IsNull()
             ? nullptr
             : Format::Unwrap(value.As<Napi::Object>())->format;
}

Napi::Value Format::NewInstance(Napi::Env env, lxw_format* format) {
  return env.GetInstanceData<Napi::ObjectReference>()
      ->Get("FormatConstructor")
      .As<Napi::Function>()
      .New({Napi::External<lxw_format>::New(env, format)});
}

Napi::Value Format::SetBgColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_bg_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetFgColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_fg_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetBorderColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_border_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetFontColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_font_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetBold(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_bold(format);
  return env.Undefined();
}

Napi::Value Format::SetBorder(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_border(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetNumFormat(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_num_format(format, info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}
