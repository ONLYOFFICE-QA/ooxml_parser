require 'spec_helper'

describe 'OoxmlSize' do
  it 'Twips and dxa units should be the same' do
    expect(OoxmlParser::OoxmlSize.new(100, :dxa)).to eq(OoxmlParser::OoxmlSize.new(100, :twip))
  end
end
