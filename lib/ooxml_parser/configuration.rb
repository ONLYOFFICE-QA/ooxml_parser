module OoxmlParser
  class Configuration
    # @return [Integer] accuracy of digits in fraction part
    attr_accessor :accuracy

    def initialize
      @accuracy = 2
    end
  end
end
