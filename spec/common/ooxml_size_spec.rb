require 'spec_helper'

describe 'OoxmlSize' do
  it 'Twips and dxa units should be the same' do
    expect(OoxmlParser::OoxmlSize.new(100, :dxa)).to eq(OoxmlParser::OoxmlSize.new(100, :twip))
  end

  describe 'one_100th_point' do
    it '1 point converts to one_100th_point' do
      expect(OoxmlParser::OoxmlSize.new(1, :point).to_unit(:one_100th_point)).to eq(OoxmlParser::OoxmlSize.new(100, :one_100th_point))
    end
  end
end
