require 'spec_helper'

describe OoxmlParser::Color do
  it 'Color looks like nil' do
    expect(OoxmlParser::Color.new(nil, nil, nil).looks_like?(nil)).to be_truthy
  end

  it 'Color with style nil looks like color nil' do
    color = OoxmlParser::Color.new(0, 0, 0)
    color.style = :nil
    expect(OoxmlParser::Color.new.looks_like?(color)).to be_truthy
  end
end
