require 'spec_helper'

describe 'OoxmlParser::Coordinates' do
  it 'Coordinates.parser_coordinates_range can handle single cell' do
    coordinates = OoxmlParser::Coordinates.parser_coordinates_range('Sheet1!$B$1')
    expect(coordinates.first).to eq(OoxmlParser::Coordinates.new(1, 'B'))
  end

  it 'Coordinates.parser_coordinates_range can handle several ranges' do
    coordinates = OoxmlParser::Coordinates.parser_coordinates_range('Donut!A7:A7,Donut!A16:A16')
    expect(coordinates.first.first).to eq(OoxmlParser::Coordinates.new(7, 'A'))
  end
end
