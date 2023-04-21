# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::OOXMLDocumentObject, '#boolean_attribute_value' do
  let(:ooxml_object) { described_class.new }

  it "returns true when given '1'" do
    expect(ooxml_object.boolean_attribute_value('1')).to be(true)
  end

  it "returns true when given 'true'" do
    expect(ooxml_object.boolean_attribute_value('true')).to be(true)
  end

  it "returns false when given '0'" do
    expect(ooxml_object.boolean_attribute_value('0')).to be(false)
  end

  it "returns false when given 'false'" do
    expect(ooxml_object.boolean_attribute_value('false')).to be(false)
  end

  it 'returns nil when given an invalid value' do
    invalid_argument = 'invalid'
    expect { ooxml_object.boolean_attribute_value(invalid_argument) }
      .to raise_error(ArgumentError,
                      "Invalid value for boolean attribute: #{invalid_argument}")
  end
end
