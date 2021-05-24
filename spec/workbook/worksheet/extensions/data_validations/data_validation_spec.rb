# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DataValidation do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/data_validations/data_validation.xlsx') }
  let(:validation) { xlsx.worksheets.first.extension_list[0].data_validations[0] }

  it 'Has uid' do
    expect(validation.uid).to eq('{0000003F-00FC-43E5-8AB4-0025003600E1}')
  end
end
