# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'body_properties_margins_values' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/'\
                                     'runs/alternate_content/office2010_content/'\
                                     'graphic/data/body_properties/margins/'\
                                     'body_properties_margins_values.docx')
    expect(docx.element_by_description[0].character_style_array[0]
               .alternate_content
               .office2010_content.graphic.data
               .body_properties.margins.top)
      .to eq(OoxmlParser::OoxmlSize.new(1.5, :centimeter))
  end
end
