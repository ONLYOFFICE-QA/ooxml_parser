# frozen_string_literal: true

require 'spec_helper'

describe 'AlphaModFix' do
  it 'alpha_mod_fix_value' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/'\
                                     'runs/alternate_content/office2010_content/'\
                                     'graphic/data/blip_fill/blip/alpha_mod_fix/'\
                                     'alpha_mod_fix_value.docx')
    expect(docx.element_by_description[0]
               .character_style_array[0].alternate_content
               .office2010_content.graphic.data
               .properties.blip_fill.blip
               .alpha_mod_fix.amount)
      .to eq(OoxmlParser::OoxmlSize.new(44, :percent))
  end
end
