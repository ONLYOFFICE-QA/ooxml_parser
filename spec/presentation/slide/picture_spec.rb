require 'spec_helper'

describe 'My behaviour' do
  it 'picture_exists.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/picture/picture_exists.pptx')
    expect(pptx.slides.first.elements.size).to be >= 3
    expect(pptx.slides.first.elements.last).to be_a_kind_of OoxmlParser::Picture
    expect(File.size?(pptx.slides.first.elements.last.image.path)).to be > 0
  end

  it 'picture_not_exist' do
    expect do
      OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/picture/picture_not_exist.pptx')
    end.to output(/Couldn't find.* file on filesystem. Possible problem in original document/).to_stderr
  end
end
