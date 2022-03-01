# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DataValidations do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/data_validations/data_validation.xlsx') }
  let(:validations) { xlsx.worksheets.first.extension_list[0].data_validations }

  it 'Got count property' do
    expect(validations.count).to eq(1)
  end

  it 'Got disable_promts value' do
    expect(validations.disable_prompts).to be(false)
  end

  it 'Contains list of array of DataValidation' do
    expect(validations[0]).to be_a(OoxmlParser::DataValidation)
  end
end
