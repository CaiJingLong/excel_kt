# ExcelUtils

Wrapper poi of apache to handle excel.

![Maven Central](https://img.shields.io/maven-central/v/top.kikt/excel)

## Include

[show version list](https://search.maven.org/artifact/top.kikt/excel)

```groovy

dependencies {
    implmentation("top.kikt:excel:$version")
}

```

### snapshot version

If you want to use develop version or snapshot version, you can add the following repository.

Usually, this will be released before Maven Central.

```groovy
repositories {
    maven {
        url "https://s01.oss.sonatype.org/content/groups/staging/"
    }
}

dependencies {
    implmentation("top.kikt:excel:$version-SNAPSHOT")
}

```

## LICENSE

Apache 2.0