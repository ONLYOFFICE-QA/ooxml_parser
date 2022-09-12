# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Coordinates, '.parser_coordinates_range' do
  it 'Coordinates.parser_coordinates_range can handle single cell' do
    coordinates = described_class.parser_coordinates_range('Sheet1!$B$1')
    expect(coordinates.first).to eq(described_class.new(1, 'B'))
  end

  it 'Coordinates.parser_coordinates_range fail with [' do
    expect(described_class.parser_coordinates_range('Sheet1!$[$1')[0].column).not_to eq('[')
  end

  it 'Coordinates.parser_coordinates_range can handle several ranges' do
    coordinates = described_class.parser_coordinates_range('Donut!A7:A7,Donut!A16:A16')
    expect(coordinates.first.first).to eq(described_class.new(7, 'A'))
  end

  it 'Coordinates.parser_coordinates_range cannot handle complex formula' do
    expect do
      described_class.parser_coordinates_range("'file!'file:xls'#'Ecc-C2345'.AJ23:'file:xls'#'Ecc-C2345'.AJ236")
    end.to output(/Formulas with # is unsupported/).to_stderr
  end

  it 'Coordinates.parser_coordinates_range not crash if formula just !' do
    expect do
      described_class.parser_coordinates_range('!')
    end.to output(/Formulas consists from `!` only/).to_stderr
  end

  it 'Coordinates.parser_coordinates_range not crash in formulas with alot of !' do
    coordinates = described_class.parser_coordinates_range('!!!!Сроки!$A$1:$E$1')
    expect(coordinates.length).to eq(5)
  end
end
