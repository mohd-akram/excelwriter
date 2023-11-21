#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"

Napi::Object Chart::Init(Napi::Env env, Napi::Object exports) {
  auto func = DefineClass(
      env,
      "Chart",
      {InstanceMethod<&Chart::AddSeries>("addSeries", napi_default_method),
       InstanceMethod<&Chart::SetTitleName>("setTitleName",
                                            napi_default_method),
       InstanceMethod<&Chart::SetTitleNameFont>("setTitleNameFont",
                                                napi_default_method),
       StaticValue("AREA_CHART",
                   Napi::Number::New(env, LXW_CHART_AREA),
                   napi_enumerable),
       StaticValue("AREA_STACKED_CHART",
                   Napi::Number::New(env, LXW_CHART_AREA_STACKED),
                   napi_enumerable),
       StaticValue("AREA_STACKED_PERCENT_CHART",
                   Napi::Number::New(env, LXW_CHART_AREA_STACKED_PERCENT),
                   napi_enumerable),
       StaticValue("BAR_CHART",
                   Napi::Number::New(env, LXW_CHART_BAR),
                   napi_enumerable),
       StaticValue("BAR_STACKED_CHART",
                   Napi::Number::New(env, LXW_CHART_BAR_STACKED),
                   napi_enumerable),
       StaticValue("BAR_STACKED_PERCENT_CHART",
                   Napi::Number::New(env, LXW_CHART_BAR_STACKED_PERCENT),
                   napi_enumerable),
       StaticValue("COLUMN_CHART",
                   Napi::Number::New(env, LXW_CHART_COLUMN),
                   napi_enumerable),
       StaticValue("COLUMN_STACKED_CHART",
                   Napi::Number::New(env, LXW_CHART_COLUMN_STACKED),
                   napi_enumerable),
       StaticValue("COLUMN_STACKED_PERCENT_CHART",
                   Napi::Number::New(env, LXW_CHART_COLUMN_STACKED_PERCENT),
                   napi_enumerable),
       StaticValue("DOUGHNUT_CHART",
                   Napi::Number::New(env, LXW_CHART_DOUGHNUT),
                   napi_enumerable),
       StaticValue("LINE_CHART",
                   Napi::Number::New(env, LXW_CHART_LINE),
                   napi_enumerable),
       StaticValue("LINE_STACKED_CHART",
                   Napi::Number::New(env, LXW_CHART_LINE_STACKED),
                   napi_enumerable),
       StaticValue("LINE_STACKED_PERCENT_CHART",
                   Napi::Number::New(env, LXW_CHART_LINE_STACKED_PERCENT),
                   napi_enumerable),
       StaticValue("PIE_CHART",
                   Napi::Number::New(env, LXW_CHART_PIE),
                   napi_enumerable),
       StaticValue("SCATTER_CHART",
                   Napi::Number::New(env, LXW_CHART_SCATTER),
                   napi_enumerable),
       StaticValue("SCATTER_STRAIGHT_CHART",
                   Napi::Number::New(env, LXW_CHART_SCATTER_STRAIGHT),
                   napi_enumerable),
       StaticValue(
           "SCATTER_STRAIGHT_WITH_MARKERS_CHART",
           Napi::Number::New(env, LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS),
           napi_enumerable),
       StaticValue("SCATTER_SMOOTH_CHART",
                   Napi::Number::New(env, LXW_CHART_SCATTER_SMOOTH),
                   napi_enumerable),
       StaticValue(
           "SCATTER_SMOOTH_WITH_MARKERS_CHART",
           Napi::Number::New(env, LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS),
           napi_enumerable),
       StaticValue("RADAR_CHART",
                   Napi::Number::New(env, LXW_CHART_RADAR),
                   napi_enumerable),
       StaticValue("RADAR_WITH_MARKERS_CHART",
                   Napi::Number::New(env, LXW_CHART_RADAR_WITH_MARKERS),
                   napi_enumerable),
       StaticValue("RADAR_FILLED_CHART",
                   Napi::Number::New(env, LXW_CHART_RADAR_FILLED),
                   napi_enumerable)

      });

  auto data = env.GetInstanceData<Napi::ObjectReference>();

  if (!data) {
    data = new Napi::ObjectReference();
    *data = Napi::Persistent(Napi::Object::New(env));
    env.SetInstanceData(data);
  }

  data->Set("ChartConstructor", func);
  exports["Chart"] = func;

  return exports;
}

Chart::Chart(const Napi::CallbackInfo& info) : Napi::ObjectWrap<Chart>(info) {
  chart = info[0].As<Napi::External<lxw_chart>>().Data();
}

lxw_chart* Chart::Get(Napi::Value value) {
  return Chart::Unwrap(value.As<Napi::Object>())->chart;
}

Napi::Value Chart::New(Napi::Env env, lxw_chart* chart) {
  return env.GetInstanceData<Napi::ObjectReference>()
      ->Get("ChartConstructor")
      .As<Napi::Function>()
      .New({Napi::External<lxw_chart>::New(env, chart)});
}

Napi::Value Chart::AddSeries(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  chart_add_series(chart,
                   info[0].IsNull()
                       ? nullptr
                       : info[0].As<Napi::String>().Utf8Value().c_str(),
                   info[1].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Chart::SetTitleName(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  chart_title_set_name(chart, info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Chart::SetTitleNameFont(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  auto options = info[0].As<Napi::Object>();
  auto name = options.Get("name");
  auto size = options.Get("size");
  auto bold = options.Get("bold");
  auto italic = options.Get("italic");
  auto underline = options.Get("underline");
  auto rotation = options.Get("rotation");
  auto color = options.Get("color");
  auto pitch_family = options.Get("pitchFamily");
  auto charset = options.Get("charset");
  auto baseline = options.Get("baseline");
  lxw_chart_font font = {};
  if (!(name.IsUndefined() || name.IsNull()))
    font.name = const_cast<char*>(name.As<Napi::String>().Utf8Value().c_str());
  if (!(size.IsUndefined() || size.IsNull()))
    font.size = size.As<Napi::Number>();
  if (!(bold.IsUndefined() || bold.IsNull()))
    font.bold = bold.As<Napi::Boolean>() ? LXW_TRUE : LXW_EXPLICIT_FALSE;
  if (!(italic.IsUndefined() || italic.IsNull()))
    font.italic = italic.As<Napi::Boolean>();
  if (!(underline.IsUndefined() || underline.IsNull()))
    font.underline = underline.As<Napi::Boolean>();
  if (!(rotation.IsUndefined() || rotation.IsNull()))
    font.rotation = rotation.As<Napi::Number>();
  if (!(color.IsUndefined() || color.IsNull()))
    font.color = color.As<Napi::Number>();
  if (!(pitch_family.IsUndefined() || pitch_family.IsNull()))
    font.pitch_family = pitch_family.As<Napi::Number>().Uint32Value();
  if (!(charset.IsUndefined() || charset.IsNull()))
    font.charset = charset.As<Napi::Number>().Uint32Value();
  if (!(baseline.IsUndefined() || baseline.IsNull()))
    font.baseline = baseline.As<Napi::Number>().Int32Value();
  chart_title_set_name_font(chart, &font);
  return env.Undefined();
}
