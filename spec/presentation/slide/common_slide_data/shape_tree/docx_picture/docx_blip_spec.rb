require 'spec_helper'

describe 'Connection Shape' do
  it 'docx_plib_file_reference_crash.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/common_slide_data/shape_tree/docx_picture/docx_blip/docx_blip_file_reference_empty.pptx')
    expect(pptx).to be_with_data
  end
end
