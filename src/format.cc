#include <napi.h>
#include <xlsxwriter.h>

#include "format.h"

Napi::Object Format::Init(Napi::Env env, Napi::Object exports) {
  auto func = DefineClass(
      env,
      "Format",
      {InstanceMethod<&Format::SetAlign>("setAlign", napi_default_method),
       InstanceMethod<&Format::SetRotation>("setRotation", napi_default_method),
       InstanceMethod<&Format::SetTextWrap>("setTextWrap", napi_default_method),
       InstanceMethod<&Format::SetIndent>("setIndent", napi_default_method),
       InstanceMethod<&Format::SetShrink>("setShrink", napi_default_method),
       InstanceMethod<&Format::SetBgColor>("setBgColor", napi_default_method),
       InstanceMethod<&Format::SetFgColor>("setFgColor", napi_default_method),
       InstanceMethod<&Format::SetBorderColor>("setBorderColor",
                                               napi_default_method),
       InstanceMethod<&Format::SetBottomColor>("setBottomColor",
                                               napi_default_method),
       InstanceMethod<&Format::SetTopColor>("setTopColor", napi_default_method),
       InstanceMethod<&Format::SetLeftColor>("setLeftColor",
                                             napi_default_method),
       InstanceMethod<&Format::SetRightColor>("setRightColor",
                                              napi_default_method),
       InstanceMethod<&Format::SetFontColor>("setFontColor",
                                             napi_default_method),
       InstanceMethod<&Format::SetFontName>("setFontName", napi_default_method),
       InstanceMethod<&Format::SetFontScript>("setFontScript",
                                              napi_default_method),
       InstanceMethod<&Format::SetFontSize>("setFontSize", napi_default_method),
       InstanceMethod<&Format::SetFontStrikeout>("setFontStrikeout",
                                                 napi_default_method),
       InstanceMethod<&Format::SetBold>("setBold", napi_default_method),
       InstanceMethod<&Format::SetItalic>("setItalic", napi_default_method),
       InstanceMethod<&Format::SetUnderline>("setUnderline",
                                             napi_default_method),
       InstanceMethod<&Format::SetBorder>("setBorder", napi_default_method),
       InstanceMethod<&Format::SetBottom>("setBottom", napi_default_method),
       InstanceMethod<&Format::SetTop>("setTop", napi_default_method),
       InstanceMethod<&Format::SetLeft>("setLeft", napi_default_method),
       InstanceMethod<&Format::SetRight>("setRight", napi_default_method),
       InstanceMethod<&Format::SetNumFormat>("setNumFormat",
                                             napi_default_method),

       // Alignment
       StaticValue("NONE_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_NONE),
                   napi_enumerable),
       StaticValue("LEFT_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_LEFT),
                   napi_enumerable),
       StaticValue("CENTER_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_CENTER),
                   napi_enumerable),
       StaticValue("RIGHT_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_RIGHT),
                   napi_enumerable),
       StaticValue("FILL_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_FILL),
                   napi_enumerable),
       StaticValue("JUSTIFY_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_JUSTIFY),
                   napi_enumerable),
       StaticValue("CENTER_ACROSS_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_CENTER_ACROSS),
                   napi_enumerable),
       StaticValue("DISTRIBUTED_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_DISTRIBUTED),
                   napi_enumerable),
       StaticValue("VERTICAL_TOP_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_VERTICAL_TOP),
                   napi_enumerable),
       StaticValue("VERTICAL_BOTTOM_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_VERTICAL_BOTTOM),
                   napi_enumerable),
       StaticValue("VERTICAL_CENTER_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_VERTICAL_CENTER),
                   napi_enumerable),
       StaticValue("VERTICAL_JUSTIFY_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_VERTICAL_JUSTIFY),
                   napi_enumerable),
       StaticValue("VERTICAL_DISTRIBUTED_ALIGN",
                   Napi::Number::New(env, LXW_ALIGN_VERTICAL_DISTRIBUTED),
                   napi_enumerable),

       // Border
       StaticValue("NONE_BORDER",
                   Napi::Number::New(env, LXW_BORDER_NONE),
                   napi_enumerable),
       StaticValue("THIN_BORDER",
                   Napi::Number::New(env, LXW_BORDER_THIN),
                   napi_enumerable),
       StaticValue("MEDIUM_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM),
                   napi_enumerable),
       StaticValue("DASHED_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DASHED),
                   napi_enumerable),
       StaticValue("DOTTED_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DOTTED),
                   napi_enumerable),
       StaticValue("THICK_BORDER",
                   Napi::Number::New(env, LXW_BORDER_THICK),
                   napi_enumerable),
       StaticValue("DOUBLE_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DOUBLE),
                   napi_enumerable),
       StaticValue("HAIR_BORDER",
                   Napi::Number::New(env, LXW_BORDER_HAIR),
                   napi_enumerable),
       StaticValue("MEDIUM_DASHED_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM_DASHED),
                   napi_enumerable),
       StaticValue("DASH_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DASH_DOT),
                   napi_enumerable),
       StaticValue("MEDIUM_DASH_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM_DASH_DOT),
                   napi_enumerable),
       StaticValue("DASH_DOT_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_DASH_DOT_DOT),
                   napi_enumerable),
       StaticValue("MEDIUM_DASH_DOT_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_MEDIUM_DASH_DOT_DOT),
                   napi_enumerable),
       StaticValue("SLANT_DASH_DOT_BORDER",
                   Napi::Number::New(env, LXW_BORDER_SLANT_DASH_DOT),
                   napi_enumerable),

       // Script
       StaticValue("SUPERSCRIPT_FONT",
                   Napi::Number::New(env, LXW_FONT_SUPERSCRIPT),
                   napi_enumerable),
       StaticValue("SUBSCRIPT_FONT",
                   Napi::Number::New(env, LXW_FONT_SUBSCRIPT),
                   napi_enumerable),

       // Underline
       StaticValue("NONE_UNDERLINE",
                   Napi::Number::New(env, LXW_UNDERLINE_NONE),
                   napi_enumerable),
       StaticValue("SINGLE_UNDERLINE",
                   Napi::Number::New(env, LXW_UNDERLINE_SINGLE),
                   napi_enumerable),
       StaticValue("DOUBLE_UNDERLINE",
                   Napi::Number::New(env, LXW_UNDERLINE_DOUBLE),
                   napi_enumerable),
       StaticValue("SINGLE_ACCOUNTING_UNDERLINE",
                   Napi::Number::New(env, LXW_UNDERLINE_SINGLE_ACCOUNTING),
                   napi_enumerable),
       StaticValue("DOUBLE_ACCOUNTING_UNDERLINE",
                   Napi::Number::New(env, LXW_UNDERLINE_DOUBLE_ACCOUNTING),
                   napi_enumerable)});

  auto data = env.GetInstanceData<Napi::ObjectReference>();

  if (!data) {
    data = new Napi::ObjectReference();
    *data = Napi::Persistent(Napi::Object::New(env));
    env.SetInstanceData(data);
  }

  data->Set("FormatConstructor", func);
  exports["Format"] = func;

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

Napi::Value Format::New(Napi::Env env, lxw_format* format) {
  return env.GetInstanceData<Napi::ObjectReference>()
      ->Get("FormatConstructor")
      .As<Napi::Function>()
      .New({Napi::External<lxw_format>::New(env, format)});
}

Napi::Value Format::SetAlign(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_align(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetRotation(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_rotation(format, info[0].As<Napi::Number>().Int32Value());
  return env.Undefined();
}

Napi::Value Format::SetTextWrap(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_text_wrap(format);
  return env.Undefined();
}

Napi::Value Format::SetIndent(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_indent(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetShrink(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_shrink(format);
  return env.Undefined();
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

Napi::Value Format::SetBottomColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_bottom_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetTopColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_top_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetLeftColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_left_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetRightColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_right_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetFontColor(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_font_color(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetFontName(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_font_name(format, info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Format::SetFontScript(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_font_script(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetFontSize(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_font_size(format, info[0].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Format::SetFontStrikeout(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_font_strikeout(format);
  return env.Undefined();
}

Napi::Value Format::SetBold(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_bold(format);
  return env.Undefined();
}

Napi::Value Format::SetItalic(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_italic(format);
  return env.Undefined();
}

Napi::Value Format::SetUnderline(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_underline(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetBorder(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_border(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetBottom(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_bottom(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetTop(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_top(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetLeft(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_left(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetRight(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_right(format, info[0].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Format::SetNumFormat(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  format_set_num_format(format, info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}
