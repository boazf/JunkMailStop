#pragma once

class CCritSect
{
public:
    CCritSect(void)
    {
        InitializeCriticalSection(&m_cs);
    }
    ~CCritSect(void)
    {
        DeleteCriticalSection(&m_cs);
    }
    void Lock(void)
    {
        EnterCriticalSection(&m_cs);
    }
    void Unlock(void)
    {
        LeaveCriticalSection(&m_cs);
    }

protected:
    CRITICAL_SECTION m_cs;
};

class CLock
{
public:
    CLock(CCritSect *pCs)
    {
        m_pcs = pCs;
        m_pcs->Lock();
    }
    ~CLock()
    {
        m_pcs->Unlock();
    }

private:
    CCritSect *m_pcs;
};

