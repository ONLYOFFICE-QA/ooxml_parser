require 'spec_helper'

describe OoxmlParser::Color do
  it 'Color looks like nil' do
    expect(OoxmlParser::Color.new(nil, nil, nil).looks_like?(nil)).to be_truthy
  end
end
