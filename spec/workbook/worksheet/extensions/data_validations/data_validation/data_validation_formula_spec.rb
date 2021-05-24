# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DataValidation do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/data_validations/data_validation.xlsx') }
  let(:validation) { xlsx.worksheets.first.extension_list[0].data_validations[0] }

  it 'Formula1 contains some formula data' do
    expect(validation.formula1.formula.value).to eq('1')
  end

  it 'Formula2 contains some formula data' do
    expect(validation.formula2.formula.value).to eq('3')
  end

  it 'Reference_sequence contains data' do
    expect(validation.reference_sequence).to eq('A1')
  end
end
