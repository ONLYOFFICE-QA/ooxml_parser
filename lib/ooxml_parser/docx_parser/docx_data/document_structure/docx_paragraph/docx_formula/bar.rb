module OoxmlParser
  class Bar
    attr_accessor :position, :element

    def initialize(position = 'bottom', element = nil)
      @position = position
      @element = element
    end
  end
end
