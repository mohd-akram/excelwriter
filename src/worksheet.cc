#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "format.h"
#include "worksheet.h"

class DataValidation {
 public:
  DataValidation(const Napi::Value& value);
  inline operator lxw_data_validation*() { return &data_validation; };

 private:
  std::string value_formula;
  std::string minimum_formula;
  std::string maximum_formula;
  std::string input_title;
  std::string input_message;
  std::string error_title;
  std::string error_message;
  std::vector<std::string> values;
  std::unique_ptr<const char*[]> value_list = nullptr;

  lxw_data_validation data_validation = {};
};

Napi::Object Worksheet::Init(Napi::Env env, Napi::Object exports) {
  auto func = DefineClass(
      env,
      "Worksheet",
      {
          InstanceMethod<&Worksheet::DataValidationCell>("dataValidationCell",
                                                         napi_default_method),
          InstanceMethod<&Worksheet::DataValidationRange>("dataValidationRange",
                                                          napi_default_method),
          InstanceMethod<&Worksheet::FreezePanes>("freezePanes",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::SplitPanes>("splitPanes",
                                                 napi_default_method),
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

          StaticValue("NONE_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_NONE),
                      napi_enumerable),
          StaticValue("INTEGER_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_INTEGER),
                      napi_enumerable),
          StaticValue(
              "INTEGER_FORMULA_VALIDATION_TYPE",
              Napi::Number::New(env, LXW_VALIDATION_TYPE_INTEGER_FORMULA),
              napi_enumerable),
          StaticValue("DECIMAL_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_DECIMAL),
                      napi_enumerable),
          StaticValue(
              "DECIMAL_FORMULA_VALIDATION_TYPE",
              Napi::Number::New(env, LXW_VALIDATION_TYPE_DECIMAL_FORMULA),
              napi_enumerable),
          StaticValue("LIST_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_LIST),
                      napi_enumerable),
          StaticValue("LIST_FORMULA_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_LIST_FORMULA),
                      napi_enumerable),
          StaticValue("DATE_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_DATE),
                      napi_enumerable),
          StaticValue("DATE_FORMULA_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_DATE_FORMULA),
                      napi_enumerable),
          StaticValue("DATE_NUMBER_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_DATE_NUMBER),
                      napi_enumerable),
          StaticValue("TIME_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_TIME),
                      napi_enumerable),
          StaticValue("TIME_FORMULA_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_TIME_FORMULA),
                      napi_enumerable),
          StaticValue("TIME_NUMBER_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_TIME_NUMBER),
                      napi_enumerable),
          StaticValue("LENGTH_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_LENGTH),
                      napi_enumerable),
          StaticValue(
              "LENGTH_FORMULA_VALIDATION_TYPE",
              Napi::Number::New(env, LXW_VALIDATION_TYPE_LENGTH_FORMULA),
              napi_enumerable),
          StaticValue(
              "CUSTOM_FORMULA_VALIDATION_TYPE",
              Napi::Number::New(env, LXW_VALIDATION_TYPE_CUSTOM_FORMULA),
              napi_enumerable),
          StaticValue("ANY_VALIDATION_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_TYPE_ANY),
                      napi_enumerable),

          StaticValue("NONE_VALIDATION_CRITERIA",
                      Napi::Number::New(env, LXW_VALIDATION_CRITERIA_NONE),
                      napi_enumerable),
          StaticValue("BETWEEN_VALIDATION_CRITERIA",
                      Napi::Number::New(env, LXW_VALIDATION_CRITERIA_BETWEEN),
                      napi_enumerable),
          StaticValue(
              "NOT_BETWEEN_VALIDATION_CRITERIA",
              Napi::Number::New(env, LXW_VALIDATION_CRITERIA_NOT_BETWEEN),
              napi_enumerable),
          StaticValue("EQUAL_TO_VALIDATION_CRITERIA",
                      Napi::Number::New(env, LXW_VALIDATION_CRITERIA_EQUAL_TO),
                      napi_enumerable),
          StaticValue(
              "NOT_EQUAL_TO_VALIDATION_CRITERIA",
              Napi::Number::New(env, LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO),
              napi_enumerable),
          StaticValue(
              "GREATER_THAN_VALIDATION_CRITERIA",
              Napi::Number::New(env, LXW_VALIDATION_CRITERIA_GREATER_THAN),
              napi_enumerable),
          StaticValue("LESS_THAN_VALIDATION_CRITERIA",
                      Napi::Number::New(env, LXW_VALIDATION_CRITERIA_LESS_THAN),
                      napi_enumerable),
          StaticValue(
              "GREATER_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA",
              Napi::Number::New(
                  env, LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO),
              napi_enumerable),
          StaticValue("LESS_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA",
                      Napi::Number::New(
                          env, LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO),
                      napi_enumerable),

          StaticValue("STOP_VALIDATION_ERROR_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_ERROR_TYPE_STOP),
                      napi_enumerable),
          StaticValue("WARNING_VALIDATION_ERROR_TYPE",
                      Napi::Number::New(env, LXW_VALIDATION_ERROR_TYPE_WARNING),
                      napi_enumerable),
          StaticValue(
              "INFORMATION_VALIDATION_ERROR_TYPE",
              Napi::Number::New(env, LXW_VALIDATION_ERROR_TYPE_INFORMATION),
              napi_enumerable),
      });

  auto data = env.GetInstanceData<Napi::ObjectReference>();

  if (!data) {
    data = new Napi::ObjectReference();
    *data = Napi::Persistent(Napi::Object::New(env));
    env.SetInstanceData(data);
  }

  data->Set("WorksheetConstructor", func);
  exports["Worksheet"] = func;

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

Napi::Value Worksheet::DataValidationCell(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_data_validation_cell(worksheet,
                                 info[0].As<Napi::Number>(),
                                 info[1].As<Napi::Number>().Uint32Value(),
                                 DataValidation(info[2]));
  return env.Undefined();
}

Napi::Value Worksheet::DataValidationRange(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_data_validation_range(worksheet,
                                  info[0].As<Napi::Number>(),
                                  info[1].As<Napi::Number>().Uint32Value(),
                                  info[2].As<Napi::Number>(),
                                  info[3].As<Napi::Number>().Uint32Value(),
                                  DataValidation(info[4]));
  return env.Undefined();
}

Napi::Value Worksheet::FreezePanes(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_freeze_panes(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Worksheet::SplitPanes(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_split_panes(
      worksheet, info[0].As<Napi::Number>(), info[1].As<Napi::Number>());
  return env.Undefined();
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
  if (info.Length() >= 4) {
    lxw_row_col_options options = {};
    auto opts = info[3].As<Napi::Object>();
    auto hidden = opts.Get("hidden");
    auto level = opts.Get("level");
    auto collapsed = opts.Get("collapsed");
    if (!hidden.IsUndefined()) options.hidden = hidden.As<Napi::Boolean>();
    if (!level.IsUndefined())
      options.level = level.As<Napi::Number>().Uint32Value();
    if (!collapsed.IsUndefined())
      options.collapsed = collapsed.As<Napi::Boolean>();
    worksheet_set_row_opt(worksheet,
                          info[0].As<Napi::Number>(),
                          info[1].As<Napi::Number>(),
                          Format::Get(info[2]),
                          &options);
  } else
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

namespace {

Napi::Number callDateMethod(const Napi::Object& date, const char* method) {
  return date.Get(method)
      .As<Napi::Function>()
      .Call(date, {})
      .As<Napi::Number>();
}

lxw_datetime convertDatetime(const Napi::Object& date) {
  lxw_datetime datetime = {};
  datetime.year = callDateMethod(date, "getFullYear");
  datetime.month = callDateMethod(date, "getMonth").Int32Value() + 1;
  datetime.day = callDateMethod(date, "getDate");
  datetime.hour = callDateMethod(date, "getHours");
  datetime.min = callDateMethod(date, "getMinutes");
  datetime.sec = callDateMethod(date, "getSeconds").DoubleValue() +
                 callDateMethod(date, "getMilliseconds").DoubleValue() / 1000.0;
  return datetime;
}

}  // namespace

DataValidation::DataValidation(const Napi::Value& value) {
  auto obj = value.As<Napi::Object>();
  auto validate = obj.Get("validate");
  auto criteria = obj.Get("criteria");
  auto ignoreBlank = obj.Get("ignoreBlank");
  auto showInput = obj.Get("showInput");
  auto showError = obj.Get("showError");
  auto errorType = obj.Get("errorType");
  auto dropdown = obj.Get("dropdown");
  auto valueNumber = obj.Get("valueNumber");
  auto valueFormula = obj.Get("valueFormula");
  auto valueList = obj.Get("valueList");
  auto valueDatetime = obj.Get("valueDatetime");
  auto minimumNumber = obj.Get("minimumNumber");
  auto minimumFormula = obj.Get("minimumFormula");
  auto minimumDatetime = obj.Get("minimumDatetime");
  auto maximumNumber = obj.Get("maximumNumber");
  auto maximumFormula = obj.Get("maximumFormula");
  auto maximumDatetime = obj.Get("maximumDatetime");
  auto inputTitle = obj.Get("inputTitle");
  auto inputMessage = obj.Get("inputMessage");
  auto errorTitle = obj.Get("errorTitle");
  auto errorMessage = obj.Get("errorMessage");

  if (!validate.IsUndefined())
    data_validation.validate = validate.As<Napi::Number>().Uint32Value();
  if (!criteria.IsUndefined())
    data_validation.criteria = criteria.As<Napi::Number>().Uint32Value();
  if (!errorType.IsUndefined())
    data_validation.error_type = errorType.As<Napi::Number>().Uint32Value();

  if (!ignoreBlank.IsUndefined())
    data_validation.ignore_blank = ignoreBlank.As<Napi::Boolean>()
                                       ? LXW_VALIDATION_ON
                                       : LXW_VALIDATION_OFF;
  if (!showInput.IsUndefined())
    data_validation.show_input =
        showInput.As<Napi::Boolean>() ? LXW_VALIDATION_ON : LXW_VALIDATION_OFF;
  if (!showError.IsUndefined())
    data_validation.show_error =
        showError.As<Napi::Boolean>() ? LXW_VALIDATION_ON : LXW_VALIDATION_OFF;
  if (!dropdown.IsUndefined())
    data_validation.dropdown =
        dropdown.As<Napi::Boolean>() ? LXW_VALIDATION_ON : LXW_VALIDATION_OFF;

  if (!valueNumber.IsUndefined())
    data_validation.value_number = valueNumber.As<Napi::Number>();
  if (!minimumNumber.IsUndefined())
    data_validation.minimum_number = minimumNumber.As<Napi::Number>();
  if (!maximumNumber.IsUndefined())
    data_validation.maximum_number = maximumNumber.As<Napi::Number>();

  if (!valueDatetime.IsUndefined())
    data_validation.value_datetime =
        convertDatetime(valueDatetime.As<Napi::Object>());
  if (!minimumDatetime.IsUndefined())
    data_validation.minimum_datetime =
        convertDatetime(minimumDatetime.As<Napi::Object>());
  if (!maximumDatetime.IsUndefined())
    data_validation.maximum_datetime =
        convertDatetime(maximumDatetime.As<Napi::Object>());

  if (!valueFormula.IsUndefined()) {
    value_formula = valueFormula.As<Napi::String>();
    data_validation.value_formula = value_formula.c_str();
  }
  if (!minimumFormula.IsUndefined()) {
    minimum_formula = minimumFormula.As<Napi::String>();
    data_validation.minimum_formula = minimum_formula.c_str();
  }
  if (!maximumFormula.IsUndefined()) {
    maximum_formula = maximumFormula.As<Napi::String>();
    data_validation.maximum_formula = maximum_formula.c_str();
  }
  if (!inputTitle.IsUndefined()) {
    input_title = inputTitle.As<Napi::String>();
    data_validation.input_title = input_title.c_str();
  }
  if (!inputMessage.IsUndefined()) {
    input_message = inputMessage.As<Napi::String>();
    data_validation.input_message = input_message.c_str();
  }
  if (!errorTitle.IsUndefined()) {
    error_title = errorTitle.As<Napi::String>();
    data_validation.error_title = error_title.c_str();
  }
  if (!errorMessage.IsUndefined()) {
    error_message = errorMessage.As<Napi::String>();
    data_validation.error_message = error_message.c_str();
  }

  if (!valueList.IsUndefined()) {
    auto array = valueList.As<Napi::Array>();
    auto length = array.Length();

    for (uint32_t i = 0; i < length; i++)
      values.push_back(array.Get(i).As<Napi::String>());

    value_list = std::make_unique<const char*[]>(length + 1);
    for (uint32_t i = 0; i < length; i++) value_list[i] = values[i].c_str();
    value_list[length] = nullptr;

    data_validation.value_list = value_list.get();
  }
}
