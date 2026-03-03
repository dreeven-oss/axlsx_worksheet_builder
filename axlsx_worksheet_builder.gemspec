# frozen_string_literal: true

require_relative "lib/axlsx_worksheet_builder/version"

Gem::Specification.new do |spec|
  spec.name = "axlsx-worksheet-builder"
  spec.version = AxlsxWorksheetBuilder::VERSION
  spec.authors = ["drvn-eb"]

  spec.summary = "Axlsx Worksheet Builder"
  spec.description = "Simple axlsx worksheet generator for iterated data"

  spec.homepage = "https://github.com/dreeven-oss/axlsx_worksheet_builder"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 3.2.0"

  spec.metadata["homepage_uri"] = spec.homepage
  spec.metadata["source_code_uri"] = "https://github.com/dreeven-oss/axlsx_worksheet_builder"

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  gemspec = File.basename(__FILE__)
  spec.files = IO.popen(%w[git ls-files -z], chdir: __dir__, err: IO::NULL) do |ls|
    ls.readlines("\x0", chomp: true).reject do |f|
      (f == gemspec) ||
        f.start_with?(*%w[bin/ Gemfile .gitignore .rspec spec/ .rubocop.yml])
    end
  end
  spec.bindir = "exe"
  spec.executables = spec.files.grep(%r{\Aexe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_dependency "caxlsx"
end
