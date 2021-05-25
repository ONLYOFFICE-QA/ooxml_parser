# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DataValidation do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/data_validations/data_validation.xlsx') }
  let(:validation) { xlsx.worksheets.first.extension_list[0].data_validations[0] }

  it 'Has uid' do
    expect(validation.uid).to eq('{0000003F-00FC-43E5-8AB4-0025003600E1}')
  end

  it 'Has type' do
    expect(validation.type).to eq(:whole)
  end

  it 'Has allow_blank' do
    expect(validation.allow_blank).to be_truthy
  end

  it 'Has error_style' do
    expect(validation.error_style).to eq(:stop)
  end

  it 'Has ime_mode' do
    expect(validation.ime_mode).to eq(:noControl)
  end

  it 'Has operator' do
    expect(validation.operator).to eq(:between)
  end

  it 'Has show_dropdown' do
    expect(validation.show_dropdown).to be_falsey
  end

  it 'Has show_input_message' do
    expect(validation.show_input_message).to be_truthy
  end

  it 'Has show_error_message' do
    expect(validation.show_error_message).to be_truthy
  end

  it 'Has error_title' do
    expect(validation.error_title).to eq('Error')
  end

  it 'Has error' do
    expect(validation.error).to eq("Value doesn't match data validation rule")
  end

  it 'Has prompt' do
    expect(validation.prompt).to eq('Enter a whole number between 1 and 3')
  end

  it 'Has prompt_title' do
    expect(validation.prompt_title).to eq('Message')
  end
end
