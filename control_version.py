with open("version.txt") as f:
    version = f.readline()

def increment_version(version):
    major, minor, patch = map(int, version.split('.'))
    patch += 1
    if patch >= 10:
        patch = 0
        minor += 1
        if minor >= 10:
            minor = 0
            major += 1
    return f"{major}.{minor}.{patch}"

with open("version.txt","w") as f:
    f.write(increment_version(version))
